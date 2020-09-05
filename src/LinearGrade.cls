VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinearGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("BridgeGeometry.Alignment.Vertical")
Option Explicit
Implements IAlignmentElement
Implements IVerticalElement

' TODO: Add tests to check expected errors

Private Type TLinearGrade
    BeginLinearGrade As CurvePoint
    EndLinearGrade As CurvePoint
End Type

Private This As TLinearGrade
Private Const BeginStationErrorMessage As String = "Beginning of linear grade station is greater than end of linear grade station."
Private Const StationNotOnCurveErrorMessage As String = "Given station is not within the limits of the linear grade."
 
Private Sub Class_Initialize()
    With This
        Set .BeginLinearGrade = New CurvePoint
        Set .EndLinearGrade = New CurvePoint
    End With
End Sub

Private Sub Class_Terminate()
    With This
        Set .BeginLinearGrade = Nothing
        Set .EndLinearGrade = Nothing
    End With
End Sub

Public Sub Create(ByVal BeginLinearGrade As CurvePoint, ByVal EndLinearGrade As CurvePoint)
    
    CheckStationLimits BeginLinearGrade.Station, EndLinearGrade.Station
   
    With This
        Set .BeginLinearGrade = BeginLinearGrade
        Set .EndLinearGrade = EndLinearGrade
    End With
    
End Sub

Public Property Get BeginStationElevation() As Double
    BeginStationElevation = This.BeginLinearGrade.Elevation
End Property

Public Property Get BeginStationValue() As Double
    BeginStationValue = This.BeginLinearGrade.StationValue
End Property

Public Property Get EndStationElevation() As Double
    EndStationElevation = This.EndLinearGrade.Elevation
End Property

Public Property Get EndStationValue() As Double
    EndStationValue = This.EndLinearGrade.StationValue
End Property

Public Function ElevationAt(ByVal Station As Station) As Double
    CheckStationOnVerticalElement Station
    ElevationAt = This.BeginLinearGrade.Elevation + Me.Grade * (Station.Value - This.BeginLinearGrade.StationValue)
End Function

Public Function Grade() As Double
    Grade = This.BeginLinearGrade.SlopeTo(This.EndLinearGrade)
End Function

Private Sub CheckStationLimits(ByVal BeginStation As Station, ByVal EndStation As Station)
    If BeginStation.Value > EndStation.Value Then
        AlignmentElementException.ThrowInvalidStationLimits _
            ErrSource:="LinearGrade", _
            ErrMessage:=BeginStationErrorMessage
    End If
End Sub

Private Sub CheckStationOnVerticalElement(ByVal Station As Station)
    If AlignmentOperations.StationOnCurveElement(Me, Station) Then
        AlignmentElementException.ThrowInvalidStationLimits _
            ErrSource:="EqualTangentParabola", _
            ErrMessage:=StationNotOnCurveErrorMessage
    End If
End Sub

' IAlignmentElement Implementation
Private Property Get IAlignmentElement_BeginStationValue() As Double
    IAlignmentElement_BeginStationValue = Me.BeginStationValue
End Property

Private Property Get IAlignmentElement_EndStationValue() As Double
    IAlignmentElement_EndStationValue = Me.EndStationElevation
End Property

' IVerticalElement Implementation
Private Property Get IVerticalElement_BeginElevationValue() As Double
    IVerticalElement_BeginElevationValue = Me.BeginStationElevation
End Property

Private Property Get IVerticalElement_BeginStationValue() As Double
    IVerticalElement_BeginStationValue = Me.BeginStationValue
End Property

Private Function IVerticalElement_ElevationAt(ByVal Station As Station) As Double
    IVerticalElement_ElevationAt = Me.ElevationAt(Station)
End Function

Private Property Get IVerticalElement_EndElevationValue() As Double
    IVerticalElement_EndElevationValue = Me.EndStationElevation
End Property

Private Property Get IVerticalElement_EndStationValue() As Double
    IVerticalElement_EndStationValue = Me.EndStationValue
End Property

Private Function IVerticalElement_SlopeAt(ByVal Station As Station) As Double
    IVerticalElement_SlopeAt = Me.Grade
End Function
