Attribute VB_Name = "BoraImpact"
' Bora macros (SPEC sections 0 and 10). Named ranges only.
' RefreshAll: the zero bid and the commitment impact table, then the stored results for every key-assumption set.
' RunCommitmentImpact, on the base set:
'   1. every lever of both bids at its zero level -> store the zero bid (ZeroBid_Stored: zero bid, then bid - zero bid);
'   2. for each commitment and bid, only that lever back at the bid's value -> store item - zero bid (Impact_Stored);
'   3. every lever restored -> last row of Impact_Stored = full bid - zero bid.
' On any error every lever and the selected set are restored before the error is shown.
Option Explicit

Private Const N_IMPACT As Long = 3                          ' COD IRR, HTM IRR, cash + guarantee at the exit

Public Sub RefreshAll()
    Application.ScreenUpdating = False
    RunCommitmentImpact
    Scenarios.RunAllScenarios
    Application.ScreenUpdating = True
End Sub

Public Sub RunCommitmentImpact()
    Dim levers As Variant, zeros As Variant, live As Range, zb As Range, store As Range
    Dim nb As Long, nz As Long, i As Long, b As Long, k As Long, c As Long
    Dim orig() As Variant, full() As Variant, zero() As Variant, zl() As Variant
    Dim origSet As Variant, errNo As Long, errMsg As String
    levers = Array("Dash_Municipal", "Dash_Community", "Dash_Overbuild", "Dash_BESS", "Dash_BESS_Hours", _
                   "Dash_PB_Mult", "Dash_Partial", "Dash_Brownfield", "Dash_Shared", "Dash_EU_Turbine", _
                   "Dash_EU_Transformer")
    zeros = Array(0, 0, 0, 0, "keep", 1, 0, 0, 0, 0, 0)     ' "keep": no zero level (BESS hours stay)
    Set live = Range("ZeroBid_Live")
    Set zb = Range("ZeroBid_Stored")
    Set store = Range("Impact_Stored")
    nb = Range("Dash_Municipal").Columns.Count
    nz = live.Rows.Count
    ReDim orig(0 To UBound(levers), 1 To nb)
    For i = 0 To UBound(levers)
        For b = 1 To nb
            orig(i, b) = Range(levers(i)).Cells(1, b).Value
        Next b
    Next i
    origSet = Range("Scenario_Name").Value
    On Error GoTo Failed
    Range("Scenario_Name").Value = Range("Scenario_Names").Cells(1, 1).Value
    Application.Calculate
    ReDim full(1 To N_IMPACT, 1 To nb)
    ReDim zero(1 To N_IMPACT, 1 To nb)
    ReDim zl(1 To nz, 1 To nb)
    For b = 1 To nb
        For k = 1 To N_IMPACT
            full(k, b) = Metric(k, b)
        Next k
        For k = 1 To nz
            zl(k, b) = live.Cells(k, b).Value
        Next k
    Next b

    ' 1. zero bid
    SetLevers levers, zeros, orig, True
    Application.Calculate
    For b = 1 To nb
        For k = 1 To N_IMPACT
            zero(k, b) = Metric(k, b)
        Next k
        For k = 1 To nz
            zb.Cells(k, b).Value = live.Cells(k, b).Value
            zb.Cells(k, nb + b).Value = Delta(zl(k, b), live.Cells(k, b).Value)
        Next k
    Next b

    ' 2. each commitment alone on top of the zero bid
    For i = 0 To UBound(levers)
        For b = 1 To nb
            If VarType(zeros(i)) <> vbString Then
                Range(levers(i)).Cells(1, b).Value = orig(i, b)
                Application.Calculate
            End If
            For k = 1 To N_IMPACT
                c = (k - 1) * nb + b
                If VarType(zeros(i)) = vbString Then
                    store.Cells(i + 1, c).ClearContents
                Else
                    store.Cells(i + 1, c).Value = Delta(Metric(k, b), zero(k, b))
                End If
            Next k
            If VarType(zeros(i)) <> vbString Then Range(levers(i)).Cells(1, b).Value = zeros(i)
        Next b
    Next i

    ' 3. full bid - zero bid, then everything back
    For b = 1 To nb
        For k = 1 To N_IMPACT
            store.Cells(store.Rows.Count, (k - 1) * nb + b).Value = Delta(full(k, b), zero(k, b))
        Next k
    Next b
    SetLevers levers, zeros, orig, False
    Range("Scenario_Name").Value = origSet
    Application.Calculate
    Exit Sub
Failed:
    errNo = Err.Number
    errMsg = Err.Description
    On Error Resume Next
    SetLevers levers, zeros, orig, False
    Range("Scenario_Name").Value = origSet
    Application.Calculate
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise errNo, "RunCommitmentImpact", errMsg & " (levers and key-assumption set restored)"
End Sub

Private Sub SetLevers(levers As Variant, zeros As Variant, orig() As Variant, ByVal toZero As Boolean)
    ' Every lever of every bid at its zero level (toZero) or back at the bid's value.
    Dim i As Long, b As Long
    For i = 0 To UBound(levers)
        For b = 1 To UBound(orig, 2)
            If toZero And VarType(zeros(i)) <> vbString Then
                Range(levers(i)).Cells(1, b).Value = zeros(i)
            Else
                Range(levers(i)).Cells(1, b).Value = orig(i, b)
            End If
        Next b
    Next i
End Sub

Private Function Metric(ByVal k As Long, ByVal b As Long) As Variant
    Select Case k
        Case 1: Metric = Range("Out_COD_IRR_" & b).Value
        Case 2: Metric = Range("Out_HTM_IRR_" & b).Value
        Case 3: Metric = Range("Out_Securities_Total_" & b).Value
    End Select
End Function

Private Function Delta(ByVal a As Variant, ByVal z As Variant) As Variant
    ' a minus z; "n/a" when either side is an error or text (e.g. a failed XIRR).
    If IsError(a) Or IsError(z) Then
        Delta = "n/a"
    ElseIf VarType(a) = vbString Or VarType(z) = vbString Then
        Delta = "n/a"
    Else
        Delta = a - z
    End If
End Function
