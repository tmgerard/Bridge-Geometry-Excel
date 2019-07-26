VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StationParserException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Alignment.Dimensioning")

Option Explicit

Public Enum StationParserError
    InvalidStationFormat = vbObjectError + 200
    NonNumericStation
End Enum

Public Sub ThrowInvalidStationStringFormat(Optional ByVal ErrSource As String, Optional ByVal ErrMessage As String)
    Err.Raise Number:=StationParserError.InvalidStationFormat, _
              Source:=ErrSource, _
              Description:=ErrMessage
End Sub

Public Sub ThrowNonNumericStation(Optional ByVal ErrSource As String, Optional ByVal ErrMessage As String)
    Err.Raise Number:=StationParserError.NonNumericStation, _
              Source:=ErrSource, _
              Description:=ErrMessage
End Sub
