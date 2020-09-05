VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CurvePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("BridgeGeometry.Alignment.Dimensioning")
'@ModuleDescription("Stores values for a point of interest along a roadway alignment.")
Option Explicit

Private Type TCurvePoint
    Station As Station
    Elevation As Double
End Type

Private This As TCurvePoint

Private Sub Class_Initialize()
    Set This.Station = New Station
End Sub

Private Sub Class_Terminate()
    Set This.Station = Nothing
End Sub

Public Sub SetCurvePoint(ByVal Station As Station, ByVal Elevation As Double)
    With This
        Set .Station = Station
        .Elevation = Elevation
    End With
End Sub

Public Property Get Elevation() As Double
    Elevation = This.Elevation
End Property

Public Property Get Station() As Station
    Set Station = This.Station
End Property

Public Property Get StationValue() As Double
    StationValue = This.Station.Value
End Property

'@Description("Checks if given curve point has the same station and elevation as compared point.")
Public Function Equals(ByVal OtherCurvePoint As CurvePoint) As Boolean

    Dim Result As Boolean
    Result = False
    
    If DoubleCompare.CompareDoubleRound(This.Elevation, OtherCurvePoint.Elevation) And _
        This.Station.Equals(OtherCurvePoint.Station) Then
        Result = True
    End If
    
    Equals = Result
End Function

'@Description("Calculate slope from object point to a given curve point object")
Public Function SlopeTo(ByVal OtherCurvePoint As CurvePoint) As Double
    
    Dim DeltaY As Double
    Dim DeltaX As Double
    
    DeltaY = OtherCurvePoint.Elevation - This.Elevation
    DeltaX = Math.Abs(OtherCurvePoint.StationValue - This.Station.Value)
    
    SlopeTo = DeltaY / DeltaX

End Function
