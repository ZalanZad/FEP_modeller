Attribute VB_Name = "BoraImpact"
' Commitment impact (Bora SPEC section 10). For each bid and each tender commitment: set the lever to its
' zero level, recalculate, store live base minus the zeroed result in Impact_Stored, restore the lever.
' Impact_Base_Live holds the headline outputs, one column per metric and bid; the last row of Impact_Stored
' stores that base, and its last column pair takes the points per EUR m of RtB price PV given up.
' Named ranges only.
Option Explicit

Public Sub RunCommitmentImpact()
    Dim levers As Variant, zeros As Variant, live As Range, store As Range
    Dim nb As Long, nm As Long, nc As Long, i As Long, b As Long, k As Long, c As Long
    Dim orig As Variant, dpv As Double
    levers = Array("Dash_Municipal", "Dash_Community", "Dash_Overbuild", "Dash_BESS", "Dash_BESS_Hours", _
                   "Dash_PB_Mult", "Dash_Partial", "Dash_Brownfield", "Dash_Shared", "Dash_EU_Turbine", _
                   "Dash_EU_Transformer")
    zeros = Array(0, 0, 0, 0, "keep", 1, 0, 0, 0, 0, 0)     ' "keep": no zero level (BESS hours stay)
    Set live = Range("Impact_Base_Live")
    Set store = Range("Impact_Stored")
    nb = Range("Dash_Municipal").Columns.Count
    nc = live.Columns.Count
    nm = nc \ nb
    Application.Calculate
    Dim baseVal() As Double
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
                Range(levers(i)).Cells(1, b).Value = zeros(i)
                Application.Calculate
                For k = 0 To nm - 1
                    c = k * nb + b
                    store.Cells(i + 1, c).Value = baseVal(c) - live.Cells(1, c).Value
                Next k
                dpv = store.Cells(i + 1, nb + b).Value              ' metric 2: RtB price PV
                If dpv < -Range("Tolerance").Value Then
                    store.Cells(i + 1, nm * nb + b).Value = store.Cells(i + 1, b).Value / -dpv
                End If
                Range(levers(i)).Cells(1, b).Value = orig
                Application.Calculate
            End If
        Next b
    Next i
End Sub
