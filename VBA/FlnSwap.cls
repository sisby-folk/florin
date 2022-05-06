VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlnSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

Private pName As String
Private pSwapCell As Range
Private pSwapSet As Name
Private pMaxSwapSet As Range
Private pSwapString As String
Private pDupeSheets As Collection
Private pCValTables As Collection
Private pDoEndOrder As Boolean
Private pDoDupeSplit As Collection

Public Property Get Name() As String
    Name = pName
End Property

Public Property Get doEndOrder() As Boolean
    doEndOrder = pDoEndOrder
End Property

Public Property Let doEndOrder(ByVal inEndOrder As Boolean)
    pDoEndOrder = inEndOrder
End Property

Public Property Get swapSet() As Name
    Set swapSet = pSwapSet
End Property

Public Property Get MaxSwapSet() As Range
    Set MaxSwapSet = pMaxSwapSet
End Property

Public Property Get swapString() As String
    swapString = pSwapString
    If Not swapSet Is Nothing Then swapString = swapSet.RefersTo
End Property

Public Property Get SwapCell() As Range
    Set SwapCell = pSwapCell
End Property

Public Property Get DupeSheets() As Collection
    Set DupeSheets = pDupeSheets
End Property

Public Property Get DoDupeSplit() As Collection
    Set DoDupeSplit = pDoDupeSplit
End Property

Public Property Get CValTables() As Collection
    Set CValTables = pCValTables
End Property

Public Function swapRange() As Range
    On Error Resume Next                         ' Returns nothing if error.
    Set swapRange = swapSet.RefersToRange
    On Error GoTo 0
End Function

Public Sub Edit(ByVal inDoEndOrder As Boolean, ByVal inSwapCell As Range, ByVal inMaxSwapSet As Range, ByVal inSwapString As String)
    pDoEndOrder = inDoEndOrder
    Set pSwapSet = Nothing
    Set pSwapCell = inSwapCell
    pSwapString = inSwapString
    Set pMaxSwapSet = inMaxSwapSet
End Sub

'' ** Signature Mirrored in FlnFlorin ** ''
Public Function Instantiate(ByVal inName As String, ByVal inDoEndOrder As Boolean, ByVal inSwapCell As Range, ByVal inMaxSwapSet As Range, ByVal inSwapString As String, ByVal inSwapSet As Name, ByVal inSheets As Collection, ByVal inCValTables As Collection, ByVal inDoDupeSplit As Collection) As FlnSwap
    If inDoDupeSplit.count <> inSheets.count Then Err.Raise vbObjectError + 513, Description:="Sheet Split Toggles don't match Sheets"
    pName = inName
    pDoEndOrder = inDoEndOrder
    Set pSwapCell = inSwapCell
    pSwapString = inSwapString
    Set pSwapSet = inSwapSet
    Set pMaxSwapSet = inMaxSwapSet
    Set pDupeSheets = inSheets
    Set pCValTables = inCValTables
    Set pDoDupeSplit = inDoDupeSplit
    Set Instantiate = Me
End Function



