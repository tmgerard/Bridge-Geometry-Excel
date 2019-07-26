VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EqualTangentParabolaFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Alignment.Vertical")
Option Explicit

Public Function CreateByPoints(ByVal PVC As CurvePoint, _
    ByVal PVIElevation As Double, _
    ByVal PVT As CurvePoint) As EqualTangentParabola

    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByCurvePoints PVC, PVIElevation, PVT
    
    Set CreateByPoints = VerticalCurve

End Function

Public Function CreateByTangentIntersection(ByVal PVI As CurvePoint, _
    ByVal Length As Double, _
    ByVal GradeIn As Double, _
    ByVal GradeOut As Double)
    
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, Length, GradeIn, GradeOut
    
    Set CreateByTangentIntersection = VerticalCurve
    
End Function

