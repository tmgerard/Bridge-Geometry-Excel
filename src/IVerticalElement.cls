VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVerticalElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Alignment.Vertical")
'@ModuleDescription("Describes station limits and elevation of vertical alignment element.")
'@IgnoreModule
'@Interface
Option Explicit

Public Property Get BeginStationValue() As Double
End Property

Public Property Get BeginElevationValue() As Double
End Property

Public Property Get EndStationValue() As Double
End Property

Public Property Get EndElevationValue() As Double
End Property

Public Function ElevationAt(ByVal Station As Station) As Double
End Function

Public Function SlopeAt(ByVal Station As Station) As Double
End Function
