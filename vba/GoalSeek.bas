Attribute VB_Name = "GoalSeek"
' Goal seek by named ranges (MODELLING_STANDARDS section 7).
' For every topic listed in the named cell GoalSeek_Topics the model defines
' <Topic>_Input (purple macro cell), <Topic>_Achieved (model result) and <Topic>_Target (target value).
' Secant iteration to machine precision (so every start point gives the same answer), capped at MAX_ITER;
' success means ABS(Achieved - Target) <= Tolerance. Circularities are re-solved after every step.
' The solved value stays in <Topic>_Input.
Option Explicit

Public Const MAX_ITER As Long = 100
Private Const PRECISION As Double = 1E-13

Public Function SolveAll() As Boolean
    Dim topics() As String, i As Long, topicList As String
    topicList = CStr(Range("GoalSeek_Topics").Value)
    SolveAll = True
    If Len(Trim(topicList)) = 0 Then Exit Function
    topics = Split(topicList, ",")
    For i = LBound(topics) To UBound(topics)
        If Not Solve(Trim(topics(i))) Then SolveAll = False
    Next i
End Function

Public Function Solve(ByVal topic As String) As Boolean
    Dim x0 As Double, x1 As Double, x2 As Double, f0 As Double, f1 As Double
    Dim tol As Double, iter As Long
    tol = Range("Tolerance").Value
    x0 = Range(topic & "_Input").Value
    f0 = Residual(topic, x0)
    x1 = x0 * 1.01
    If x1 = 0 Then x1 = 1
    f1 = Residual(topic, x1)
    For iter = 1 To MAX_ITER
        If f1 = 0 Or f1 = f0 Or Abs(x1 - x0) <= Abs(x1) * PRECISION Then Exit For
        x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        x0 = x1
        f0 = f1
        x1 = x2
        f1 = Residual(topic, x1)
    Next iter
    Solve = (Abs(f1) <= tol)
End Function

Private Function Residual(ByVal topic As String, ByVal x As Double) As Double
    Range(topic & "_Input").Value = x
    Application.Calculate
    Solver.SolveAll
    Residual = Range(topic & "_Achieved").Value - Range(topic & "_Target").Value
End Function
