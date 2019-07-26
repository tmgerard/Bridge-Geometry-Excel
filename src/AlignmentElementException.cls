VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AlignmentElementException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Alignment")

Option Explicit

Public Enum AlignmentElementError
    InvalidStationLimits = vbObjectError + 300
End Enum

Public Sub ThrowInvalidStationLimits(Optional ByVal ErrSource As String, Optional ByVal ErrMessage As String)
    Err.Raise Number:=AlignmentElementError.InvalidStationLimits, _
              Source:=ErrSource, _
              Description:=ErrMessage
End Sub
