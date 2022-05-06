VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlnProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

Private pName As String
Private pRules As Collection
Private pAssetLists As Collection
Private pOutputTables As Collection
Private pSwaps As Collection
Private pPhotoFills As Collection
Private pPageGroups As Collection
Private pAutofits As Collection
Private pAutohides As Collection
Private pAssetCell As Range
Private pPhotoRename As Range
Private pPhotoPath As String
Private pFilenameCell As Range

Public Property Get autofits() As Collection
    Set autofits = pAutofits
End Property
Public Property Get autohides() As Collection
    Set autohides = pAutohides
End Property

Public Property Set autofits(ByVal RHS As Collection)
    Set pAutofits = RHS
End Property

Public Property Get pageGroups() As Collection
    Set pageGroups = pPageGroups
End Property

Public Property Set pageGroups(ByVal RHS As Collection)
    Set pPageGroups = RHS
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal inName As String)
    pName = inName
End Property

Public Property Get Rules() As Collection
    Set Rules = pRules
End Property

Public Property Get AssetLists() As Collection
    Set AssetLists = pAssetLists
End Property

Public Property Get OutputTables() As Collection
    Set OutputTables = pOutputTables
End Property

Public Property Get Swaps() As Collection
    Set Swaps = pSwaps
End Property

Public Property Get photofills() As Collection
    Set photofills = pPhotoFills
End Property

Public Property Get AssetCell() As Range
    Set AssetCell = pAssetCell
End Property
Public Property Get PhotoRename() As Range
    Set PhotoRename = pPhotoRename
End Property
Public Property Get photoPath() As String
    photoPath = pPhotoPath
End Property

Public Property Get FilenameCell() As Range
    Set FilenameCell = pFilenameCell
End Property

Public Property Set AssetCell(ByVal inAssetCell As Range)
    Set pAssetCell = inAssetCell
End Property
Public Property Set PhotoRename(ByVal inPhotoRename As Range)
    Set pPhotoRename = inPhotoRename
End Property
Public Property Let photoPath(ByVal inPhotoPath As String)
    pPhotoPath = inPhotoPath
End Property

Public Property Set FilenameCell(ByVal inFilenameCell As Range)
    Set pFilenameCell = inFilenameCell
End Property

Public Function GetSheets() As Collection
    Dim outSheets As Collection
    Set outSheets = New Collection
    Dim currentRule As FlnRule
    For Each currentRule In pRules
        collectionAppendSheets outSheets, currentRule.GetSheets
    Next
    Set GetSheets = outSheets
End Function

Public Function getActiveSheets() As Collection
    Dim outSheets As Collection
    Set outSheets = New Collection
    Dim currentRule As FlnRule
    For Each currentRule In pRules
        Dim activeCase As FlnCase
        Set activeCase = currentRule.getFirstActiveCase
        If Not activeCase Is Nothing Then collectionAppendSheets outSheets, activeCase.Sheets
    Next
    Set getActiveSheets = outSheets
End Function

Public Function getHideRange() As SvnMultiRange
    Dim outRange As SvnMultiRange
    Dim currentRule As FlnRule
    For Each currentRule In pRules
        Dim currentCase As FlnCase
        For Each currentCase In currentRule.Cases
            If Not currentCase Is Nothing Then
                If outRange Is Nothing Then Set outRange = multiRangeFromInstantiate(ThisWorkbook, inUnionize:=True, inEntireRows:=True)
                outRange.Union currentCase.HideRef
            End If
        Next
    Next
    Set getHideRange = outRange
End Function

Public Function getActiveHide() As SvnMultiRange
    Dim outRange As SvnMultiRange
    Dim currentRule As FlnRule
    For Each currentRule In pRules
        Dim currentCase As FlnCase
        Set currentCase = currentRule.getFirstActiveCase
        If Not currentCase Is Nothing Then
            If outRange Is Nothing Then Set outRange = multiRangeFromInstantiate(ThisWorkbook, inUnionize:=True, inEntireRows:=True)
            outRange.Union currentCase.HideRef
        End If
    Next
    Set getActiveHide = outRange
End Function

Public Function flagDoesExport() As Boolean
    flagDoesExport = Me.GetSheets.count > 0
End Function
Public Function flagHasPhotos() As Boolean
    flagHasPhotos = pPhotoFills.count > 0
End Function
Public Function flagRenamesPhotos() As Boolean
    flagRenamesPhotos = Not pPhotoRename Is Nothing
End Function
Public Function flagHasSwaps() As Boolean
    flagHasSwaps = pSwaps.count > 0
End Function
Public Function flagHasOutputs() As Boolean
    flagHasOutputs = pOutputTables.count > 0
End Function

'' ** Signature Mirrored in FlnFlorin ** ''
Public Function Instantiate(ByVal inName As String, ByVal inRules As Collection, ByVal inAssetLists As Collection, ByVal inOutputTables As Collection, ByVal inSwaps As Collection, ByVal inPhotoFills As Collection, ByVal inAutofits As Collection, ByVal inAutohides As Collection, ByVal inPageGroups As Collection, ByVal inAssetCell As Range, ByVal inFilenameCell As Range, ByVal inPhotoPath As String, ByVal inPhotoRename As Range) As FlnProfile
    pName = inName
    Set pRules = inRules
    Set pAssetLists = inAssetLists
    Set pOutputTables = inOutputTables
    Set pSwaps = inSwaps
    Set pPhotoFills = inPhotoFills
    Set pAutofits = inAutofits
    Set pAutohides = inAutohides
    Set pPageGroups = inPageGroups
    Set pAssetCell = inAssetCell
    Set pPhotoRename = inPhotoRename
    Set pFilenameCell = inFilenameCell
    pPhotoPath = inPhotoPath
    Set Instantiate = Me
End Function

