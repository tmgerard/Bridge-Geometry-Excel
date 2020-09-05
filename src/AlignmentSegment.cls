VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AlignmentSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("BridgeGeometry.Alignment")
Option Explicit

Private Type TAlignmentSegment
    Begin As Station
    End As Station
End Type

Private This As TAlignmentSegment

Private Sub Class_Initialize()
    With This
        Set .Begin = New Station
        Set .End = New Station
    End With
End Sub

Private Sub Class_Terminate()
    With This
        Set .Begin = Nothing
        Set .End = Nothing
    End With
End Sub

Public Property Get BeginStation() As Station
    Set BeginStation = This.Begin
End Property

Public Property Set BeginStation(ByRef BeginStation As Station)
    Set This.Begin = BeginStation
End Property

Public Property Get EndStation() As Station
    Set EndStation = This.End
End Property

Public Function Length() As Double
    Length = This.End.Value - This.Begin.Value
End Function
