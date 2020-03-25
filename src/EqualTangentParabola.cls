VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EqualTangentParabola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("BridgeGeometry.Alignment.Vertical")
Option Explicit
Implements IVerticalElement
Implements IAlignmentElement

Private Type TEqualTangentParabola
    BeginVerticalCurve As CurvePoint
    TangentIntersection As CurvePoint
    EndVerticalCurve As CurvePoint
    GradeIn As Double
    GradeOut As Double
End Type

Private this As TEqualTangentParabola
Private Const BeginStationErrorMessage As String = "Beginning of vertical curve station is greater than end of vertical curve station."
Private Const StationNotOnCurveErrorMessage As String = "Given station is not within the limits of the vertical curve."
Private Const DecimalToPercent As Double = 100

Private Sub Class_Initialize()
    With this
        Set .BeginVerticalCurve = New CurvePoint
        Set .TangentIntersection = New CurvePoint
        Set .EndVerticalCurve = New CurvePoint
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .BeginVerticalCurve = Nothing
        Set .TangentIntersection = Nothing
        Set .EndVerticalCurve = Nothing
    End With
End Sub

'@Ignore UseMeaningfulName
Public Sub CreateByCurvePoints(ByVal PVC As CurvePoint, ByVal PVIElevation As Double, ByVal PVT As CurvePoint)
    
    CheckStationLimits PVC.Station, PVT.Station
    
    Dim PVIStation As Station
    Set PVIStation = StationMath.AddValueToStation(PVC.Station, (PVT.StationValue - PVC.StationValue) / 2)
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, PVIElevation
    
    With this
        Set .BeginVerticalCurve = PVC
        Set .TangentIntersection = PVI
        Set .EndVerticalCurve = PVT
        .GradeIn = PVC.SlopeTo(PVI)
        .GradeOut = PVI.SlopeTo(PVT)
    End With
End Sub

Public Sub CreateByTangentIntersection(ByVal PVI As CurvePoint, ByVal Length As Double, _
    ByVal GradeIn As Double, ByVal GradeOut As Double)
    
    Dim PVCStation As Station
    Set PVCStation = New Station
    PVCStation.Value = PVI.StationValue - (Length / 2)
    
    Dim PVCElevation As Double
    PVCElevation = PVI.Elevation - ((GradeIn) * (Length / 2))
    
    Dim PVTStation As Station
    Set PVTStation = New Station
    PVTStation.Value = PVI.StationValue + (Length / 2)
    
    Dim PVTElevation As Double
    PVTElevation = PVI.Elevation + ((GradeOut) * (Length / 2))
    
    With this
        .BeginVerticalCurve.SetCurvePoint PVCStation, PVCElevation
        Set .TangentIntersection = PVI
        .EndVerticalCurve.SetCurvePoint PVTStation, PVTElevation
        .GradeIn = GradeIn
        .GradeOut = GradeOut
    End With
    
End Sub

Public Property Get BeginStationValue() As Double
    BeginStationValue = this.BeginVerticalCurve.StationValue
End Property

Public Property Get BeginElevationValue() As Double
    BeginElevationValue = this.BeginVerticalCurve.Elevation
End Property

Public Property Get EndStationValue() As Double
    EndStationValue = this.EndVerticalCurve.StationValue
End Property

Public Property Get EndElevationValue() As Double
    EndElevationValue = this.EndVerticalCurve.Elevation
End Property

Public Property Get GradeIn() As Double
    GradeIn = this.GradeIn
End Property

Public Property Get GradeOut() As Double
    GradeOut = this.GradeOut
End Property

Public Function ChangeInGradient() As Double
    ChangeInGradient = DecimalToPercent * Math.Abs(this.GradeOut - this.GradeIn)
End Function

Public Function Length() As Double
    Length = this.EndVerticalCurve.StationValue - this.BeginVerticalCurve.StationValue
End Function

Public Function MiddleOrdinate() As Double
    MiddleOrdinate = ChangeInGradient * (Me.Length / FeetPerStation) / 8
End Function

Public Function RateOfGradeChangePerStation() As Double
    RateOfGradeChangePerStation = DecimalToPercent * (this.GradeOut - this.GradeIn) / (Me.Length / FeetPerStation)
End Function

Public Function ToTurningPoint() As Double
    ToTurningPoint = FeetPerStation * (DecimalToPercent * -this.GradeIn / Me.RateOfGradeChangePerStation)
End Function

Public Function ElevationAt(ByVal Station As Station) As Double

    CheckStationOnVerticalElement Station

    Dim DistanceToPointInStations As Double
    DistanceToPointInStations = (Station.Value - this.BeginVerticalCurve.StationValue) / FeetPerStation
    
    '@Ignore UseMeaningfulName
    Dim Term1 As Double
    Term1 = (Me.RateOfGradeChangePerStation / 2) * DistanceToPointInStations ^ 2
    
    '@Ignore UseMeaningfulName
    Dim Term2 As Double
    Term2 = (DecimalToPercent * this.GradeIn) * DistanceToPointInStations
    
    ElevationAt = Term1 + Term2 + this.BeginVerticalCurve.Elevation
End Function

Public Function SlopeAt(ByVal Station As Station) As Double

    CheckStationOnVerticalElement Station

    Dim DistanceToPointInStations As Double
    DistanceToPointInStations = (Station.Value - this.BeginVerticalCurve.StationValue) / FeetPerStation
    
    '@Ignore UseMeaningfulName
    Dim Term1 As Double
    Term1 = Me.RateOfGradeChangePerStation * DistanceToPointInStations
    
    SlopeAt = Term1 + (DecimalToPercent * this.GradeIn)

End Function

Private Sub CheckStationLimits(ByVal BeginStation As Station, ByVal EndStation As Station)
    If BeginStation.Value > EndStation.Value Then
        AlignmentElementException.ThrowInvalidStationLimits _
            ErrSource:="EqualTangentParabola", _
            ErrMessage:=BeginStationErrorMessage
    End If
End Sub

Private Sub CheckStationOnVerticalElement(ByVal Station As Station)
    If Not AlignmentOperations.StationOnCurveElement(Me, Station) Then
        AlignmentElementException.ThrowInvalidStationLimits _
            ErrSource:="EqualTangentParabola", _
            ErrMessage:=StationNotOnCurveErrorMessage
    End If
End Sub

' IAlignment Implementation
Private Property Get IAlignmentElement_BeginStationValue() As Double
    IAlignmentElement_BeginStationValue = Me.BeginStationValue
End Property

Private Property Get IAlignmentElement_EndStationValue() As Double
    IAlignmentElement_EndStationValue = Me.EndStationValue
End Property

' IVerticalElement Implementation
Private Property Get IVerticalElement_BeginElevationValue() As Double
    IVerticalElement_BeginElevationValue = Me.BeginElevationValue
End Property

Private Property Get IVerticalElement_BeginStationValue() As Double
    IVerticalElement_BeginStationValue = Me.BeginStationValue
End Property

Private Function IVerticalElement_ElevationAt(ByVal Station As Station) As Double
    IVerticalElement_ElevationAt = Me.ElevationAt(Station)
End Function

Private Property Get IVerticalElement_EndElevationValue() As Double
    IVerticalElement_EndElevationValue = Me.EndElevationValue
End Property

Private Property Get IVerticalElement_EndStationValue() As Double
    IVerticalElement_EndStationValue = Me.EndStationValue
End Property

Private Function IVerticalElement_SlopeAt(ByVal Station As Station) As Double
    IVerticalElement_SlopeAt = Me.SlopeAt(Station)
End Function








