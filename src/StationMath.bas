Attribute VB_Name = "StationMath"
'@Folder("Alignment.Dimensioning")
Option Explicit
Option Private Module

Public Const FeetPerStation As Long = 100

Public Function AddStations(ByVal LeftStation As Station, ByVal RightStation As Station) As Station
    Dim Result As Station
    Set Result = New Station
    Result.Value = LeftStation.Value + RightStation.Value
    Set AddStations = Result
End Function

Public Function AddValueToStation(ByVal LeftStation As Station, ByVal Value As Double) As Station
    Dim Result As Station
    Set Result = New Station
    Result.Value = LeftStation.Value + Value
    Set AddValueToStation = Result
End Function

Public Function DifferenceInStations(ByVal LeftStation As Station, ByVal RightStation As Station) As Double
    DifferenceInStations = (LeftStation.Value - RightStation.Value) / FeetPerStation
End Function

Public Function SubtractStations(ByVal LeftStation As Station, ByVal RightStation As Station) As Station
    Dim Result As Station
    Set Result = New Station
    Result.Value = LeftStation.Value - RightStation.Value
    Set SubtractStations = Result
End Function

Public Function SubtractValueFromStation(ByVal LeftStation As Station, ByVal Value As Double) As Station
    Dim Result As Station
    Set Result = New Station
    Result.Value = LeftStation.Value - Value
    Set SubtractValueFromStation = Result
End Function


