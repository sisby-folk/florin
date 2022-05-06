VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlnRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

Private pName As String
Private pCases As Collection

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal inName As String)
    pName = inName
End Property

Public Property Get Cases() As Collection
    Set Cases = pCases
End Property

Public Function GetSheets() As Collection
    Dim outSheets As Collection
    Set outSheets = New Collection
    Dim currentCase As FlnCase
    For Each currentCase In pCases
        collectionAppendSheets outSheets, currentCase.Sheets
    Next
    Set GetSheets = outSheets
End Function

Public Function getFirstActiveCase() As FlnCase
    Dim currentCase As FlnCase
    Set getFirstActiveCase = Nothing
    For Each currentCase In pCases
        If currentCase.isActive Then
            Set getFirstActiveCase = currentCase
            Exit For
        End If
    Next
End Function

'' ** Signature Mirrored in FlnFlorin ** ''
Public Function Instantiate(ByVal inName As String, ByVal inCases As Collection) As FlnRule
    pName = inName
    Set pCases = inCases
    Set Instantiate = Me
End Function

