Attribute VB_Name = "BezierModule"
Option Explicit

Public Type BezierPoint
    X As Long
    Y As Long
End Type

Private Function Factorial(ByVal n As Long) As Long
    Factorial = 1
    For n = n To 1 Step -1
        Factorial = Factorial * n
    Next
End Function

Function CalculateBezier(Time As Double, ControlPoints() As BezierPoint) As BezierPoint
    Dim Pos As Long, Max As Long
    Dim Multiplier As Double, Coefficient As Long
    Max = UBound(ControlPoints)
    For Pos = 0 To Max
        Coefficient = Factorial(Max) / (Factorial(Pos) * Factorial(Max - Pos))
        Multiplier = (Time ^ Pos) * ((1 - Time) ^ (Max - Pos)) * Coefficient
        CalculateBezier.X = CalculateBezier.X + (ControlPoints(Pos).X * Multiplier)
        CalculateBezier.Y = CalculateBezier.Y + (ControlPoints(Pos).Y * Multiplier)
    Next
End Function
