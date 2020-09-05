VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Station"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("BridgeGeometry.Alignment.Dimensioning")
'@ModuleDescription("Describes linear dimension in roadway stationing format.")
Option Explicit

Private Type TStation
    StationValue As Double
End Type

Private This As TStation
Private Const DefaultStationFormat As String = "#0+00.00#"

Public Property Get Value() As Double
    Value = This.StationValue
End Property

Public Property Let Value(ByVal StationValue As Double)
    This.StationValue = StationValue
End Property

Public Sub AddStation(ByVal OtherStation As Station)
    This.StationValue = This.StationValue + OtherStation.Value
End Sub

Public Sub AddValue(ByVal Value As Double)
    This.StationValue = This.StationValue + Value
End Sub

Public Function Equals(ByVal OtherStation As Station) As Boolean
    Equals = DoubleCompare.CompareDoubleRound(This.StationValue, OtherStation.Value)
End Function

Public Sub SubtractStation(ByVal OtherStation As Station)
    This.StationValue = This.StationValue - OtherStation.Value
End Sub

Public Sub SubtractValue(ByVal Value As Double)
    This.StationValue = This.StationValue - Value
End Sub

Public Function ToString(Optional ByVal StringFormat As String = DefaultStationFormat) As String
    ToString = Format$(This.StationValue, StringFormat)
End Function

