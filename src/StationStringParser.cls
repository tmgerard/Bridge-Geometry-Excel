VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StationStringParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@PredeclaredId
'@Folder("BridgeGeometry.Alignment.Dimensioning")

Option Explicit

Private Const StationDelimeter As String = "+"
Private Const StationSplitArrayLimit As Long = 2
Private Const ArraySizeErrorMessage As String = "Invalid station formatting. Only one delimeter ('+') per string allowed."
Private Const NonNumericStationErrorMessage As String = "Station parts are not numeric."

Public Function ToDouble(ByVal StationString As String) As Double

    Dim Temp() As String
    Temp = Split(StationString, StationDelimeter)
    
    CheckSplitArraySize StationArray:=Temp
    CheckIsNumericStation StationArray:=Temp
    
    ToDouble = CDbl(Temp(0)) * FeetPerStation + CDbl(Temp(1))

End Function

'@Description("Station that split array is correct size. One delimeter will create an array with two elements.")
Private Sub CheckSplitArraySize(ByRef StationArray() As String)

    If Not UBound(StationArray) + 1 = StationSplitArrayLimit Then
        StationParserException.ThrowInvalidStationStringFormat _
            ErrSource:="StationParser.CheckSplitArraySize", _
            ErrMessage:=ArraySizeErrorMessage
    End If

End Sub

'@Description("Check that station array contains numeric elements.")
Private Sub CheckIsNumericStation(ByRef StationArray() As String)

    Dim Item As Long
    For Item = LBound(StationArray) To UBound(StationArray)
        If Not IsNumeric(StationArray(Item)) Then
            StationParserException.ThrowNonNumericStation _
                ErrSource:="StationParser.CheckIsNumericStation", _
                ErrMessage:=NonNumericStationErrorMessage
        End If
    Next Item

End Sub





