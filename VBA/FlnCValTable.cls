VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlnCValTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

Private pName As String
Private pSource As Range                         ' Single Range
Private pDest As Range                           ' Multi-Area per additional swap

Public Property Get Name() As String
    Name = pName
End Property

Public Property Get Source() As Range
    Set Source = pSource
End Property

Public Property Get Dest() As Range
    Set Dest = pDest
End Property

'' ** Signature Mirrored in FlnFlorin ** ''
Public Function Instantiate(ByVal inName As String, ByVal inSource As Range, ByVal inDest As Range) As FlnCValTable
    pName = inName
    Set pSource = inSource
    Set pDest = inDest
    Set Instantiate = Me
End Function

