Attribute VB_Name = "Solver"
' Circularity breaker: COPY -> PASTE loop (MODELLING_STANDARDS section 7).
' For every topic listed in the named cell Solver_Topics (comma-separated) the model defines
' <Topic>_Copy (live formula row), <Topic>_Paste (macro-written row) and <Topic>_Delta (sum of ABS differences).
' Named ranges only; capped at MAX_ITER iterations.
Option Explicit

Public Const MAX_ITER As Long = 100

Public Function SolveAll() As Boolean
    SolveAll = SolveTopics(CStr(Range("Solver_Topics").Value))
End Function

Public Function SolveTopics(ByVal topicList As String) As Boolean
    Dim topics() As String, i As Long, iter As Long
    If Len(Trim(topicList)) = 0 Then
        SolveTopics = True
        Exit Function
    End If
    topics = Split(topicList, ",")
    Application.Calculate
    For iter = 1 To MAX_ITER
        If DeltasAreZero(topics) Then
            SolveTopics = True
            Exit Function
        End If
        For i = LBound(topics) To UBound(topics)
            Range(Trim(topics(i)) & "_Paste").Value = Range(Trim(topics(i)) & "_Copy").Value
        Next i
        Application.Calculate
    Next iter
    SolveTopics = DeltasAreZero(topics)
End Function

Private Function DeltasAreZero(topics() As String) As Boolean
    Dim i As Long
    DeltasAreZero = True
    For i = LBound(topics) To UBound(topics)
        If Abs(Range(Trim(topics(i)) & "_Delta").Value) > Range("Tolerance").Value Then DeltasAreZero = False
    Next i
End Function
