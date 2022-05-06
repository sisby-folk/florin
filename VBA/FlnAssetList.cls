VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlnAssetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

Private pName As String
Private pRange As SvnMultiRange

Public Property Get Name() As String
    Name = pName
End Property

Public Property Get MultiRange() As SvnMultiRange
    Set MultiRange = pRange
End Property

Public Function ToString() As Variant
    ToString = pRange.ToString
End Function

Public Sub setRange(ByVal inRange As String)
    pRange.Instantiate ThisWorkbook, inRange
End Sub

'' ** Signature Mirrored in FlnFlorin ** ''
Public Function Instantiate(ByVal inName As String, ByVal inRange As String) As FlnAssetList
    pName = inName
    Set pRange = multiRangeFromInstantiate(ThisWorkbook, inRange)
    Set Instantiate = Me
End Function

