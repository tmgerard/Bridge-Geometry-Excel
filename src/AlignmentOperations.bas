Attribute VB_Name = "AlignmentOperations"
'@Folder("Alignment")
Option Explicit

Public Function StationOnCurveElement(ByVal AlignmentElement As IAlignmentElement, _
    ByVal Station As Station) As Boolean

    Dim Result As Boolean

    If Station.Value < AlignmentElement.BeginStationValue Then
        Result = False
    ElseIf Station.Value > AlignmentElement.EndStationValue Then
        Result = False
    Else
        Result = True
    End If
    
    StationOnCurveElement = Result

End Function
