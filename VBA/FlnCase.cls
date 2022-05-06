VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlnCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

Private pCondition As Variant
Private pHideRef As SvnMultiRange
Private pSheets As Collection

Public Property Get Condition() As Variant
    Condition = pCondition
End Property

Public Property Let Condition(ByVal inCondition As Variant)
    pCondition = inCondition
End Property

Public Property Get HideRef() As SvnMultiRange
    Set HideRef = pHideRef
End Property

Public Property Set HideRef(ByVal inHideRef As SvnMultiRange)
    Set pHideRef = inHideRef
End Property

Public Property Get Sheets() As Collection
    Set Sheets = pSheets
End Property

Public Function isActive() As Boolean
    Dim cond As Variant
    cond = boolOrEmpty(pCondition)
    If Not IsEmpty(cond) Then isActive = CBool(cond) Else isActive = False
End Function

'' ** Signature Mirrored in FlnFlorin ** ''
Public Function Instantiate(ByVal inCondition As Variant, ByVal inHideRef As SvnMultiRange, ByVal inSheets As Collection) As FlnCase
    pCondition = inCondition
    Set pHideRef = inHideRef
    Set pSheets = inSheets
    Set Instantiate = Me
End Function

