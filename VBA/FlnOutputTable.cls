VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlnOutputTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

Private pName As String
Private pIdList As Range                         ' Sorted IDs for sorted insert
Private pSource As Range                         ' Multi-Area: Sigs, Data, Valid
Private pDest As Range                           ' Multi-Area: IDs, Sigs, Data

Private AlgoIDRows As Collection
Private AlgoSigRows As Collection

Public Property Get Name() As String
    Name = pName
End Property

Public Property Get IdList() As Range
    Set IdList = pIdList
End Property

Public Property Get Source() As Range
    Set Source = pSource
End Property

Public Property Get Dest() As Range
    Set Dest = pDest
End Property

Public Function GetIDRow(ByVal currentValue As String) As Long
    If collectionGetOrNothing(AlgoIDRows, currentValue) Is Nothing Then
        GetIDRow = AlgoIDRows.count + 1
        Exit Function
    End If
    GetIDRow = AlgoIDRows(currentValue)
End Function

Public Function GetSigRow(ByVal currentValue As Range) As Long
    GetSigRow = AlgoSigRows(rangeGetText(currentValue))
End Function

Public Sub SetupAlgo()
    ' Row lookups for ordering
    Set AlgoIDRows = New Collection
    Set AlgoSigRows = New Collection
    
    ' TODO: Catch non-present errors
    Dim currentRow As Long
    For currentRow = 1 To pIdList.Rows.count
        AlgoIDRows.Add currentRow, vbNullString & rangeGetText(pIdList.Rows(currentRow)) 'Lookup for ordering
    Next
    For currentRow = 1 To pSource.Areas(1).Rows.count
        AlgoSigRows.Add currentRow, rangeGetText(pSource.Areas(1).Rows(currentRow)) 'Lookup for ordering
    Next
End Sub

'' ** Signature Mirrored in FlnFlorin ** ''
Public Function Instantiate(ByVal inName As String, ByVal inIdList As Range, ByVal inSource As Range, ByVal inDest As Range) As FlnOutputTable
    pName = inName
    Set pIdList = inIdList
    Set pSource = inSource
    Set pDest = inDest
    Set Instantiate = Me
End Function

