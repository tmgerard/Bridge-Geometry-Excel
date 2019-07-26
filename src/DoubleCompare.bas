Attribute VB_Name = "DoubleCompare"
'@Folder("Utility")
Option Explicit

Private Const Epsilon As Double = 2.22044604925031E-16

'@Description("Compare two double precision values using machine epsilon value (est. 2^-52)")
Public Function CompareDoubleEpsilon(ByVal ValueOne As Double, ByVal ValueTwo As Double) As Boolean
    CompareDoubleEpsilon = (Math.Abs(ValueOne - ValueTwo) < Epsilon)
End Function

'@Description("Compare two double precision values using the round function.")
Public Function CompareDoubleRound(ByVal ValueOne As Double, ByVal ValueTwo As Double, _
    Optional ByVal Precision As Long = 8) As Boolean
    CompareDoubleRound = (Math.Round(ValueOne, Precision) = Math.Round(ValueTwo, Precision))
End Function

