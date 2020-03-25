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

Private this As TStation
Private Const DefaultStationFormat As String = "#0+00.00#"

Public Property Get Value() As Double
    Value = this.StationValue
End Property

Public Property Let Value(ByVal StationValue As Double)
    this.StationValue = StationValue
End Property

Public Sub AddStation(ByVal OtherStation As Station)
    this.StationValue = this.StationValue + OtherStation.Value
End Sub

Public Sub AddValue(ByVal Value As Double)
    this.StationValue = this.StationValue + Value
End Sub

Public Function Equals(ByVal OtherStation As Station) As Boolean
    Equals = DoubleCompare.CompareDoubleRound(this.StationValue, OtherStation.Value)
End Function

Public Sub SubtractStation(ByVal OtherStation As Station)
    this.StationValue = this.StationValue - OtherStation.Value
End Sub

Public Sub SubtractValue(ByVal Value As Double)
    this.StationValue = this.StationValue - Value
End Sub

Public Function ToString(Optional ByVal StringFormat As String = DefaultStationFormat) As String
    ToString = Format$(this.StationValue, StringFormat)
End Function

