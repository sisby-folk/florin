Attribute VB_Name = "FlnFlorin"
'@Folder("Florin")
Option Explicit

'*********************************'
'' ** Freehand Instantiations ** ''
'*********************************'

Public Function assetListFromInstantiate(ByVal inName As String, ByVal inRange As String) As FlnAssetList
    Set assetListFromInstantiate = New FlnAssetList
    assetListFromInstantiate.Instantiate inName:=inName, inRange:=inRange
End Function

Public Function autoFitFromInstantiate(ByVal inName As String, ByVal inRange As String) As FlnAutoFit
    Set autoFitFromInstantiate = New FlnAutoFit
    autoFitFromInstantiate.Instantiate inName:=inName, inRange:=inRange
End Function

Public Function autoHideFromInstantiate(ByVal inName As String, ByVal inRange As String) As FlnAutoHide
    Set autoHideFromInstantiate = New FlnAutoHide
    autoHideFromInstantiate.Instantiate inName:=inName, inRange:=inRange
End Function

Public Function caseFromInstantiate(ByVal inCondition As Variant, ByVal inHideRef As SvnMultiRange, ByVal inSheets As Collection) As FlnCase
    Set caseFromInstantiate = New FlnCase
    caseFromInstantiate.Instantiate inCondition:=inCondition, inHideRef:=inHideRef, inSheets:=inSheets
End Function

Public Function cValTableFromInstantiate(ByVal inName As String, ByVal inSource As Range, ByVal inDest As Range) As FlnCValTable
    Set cValTableFromInstantiate = New FlnCValTable
    cValTableFromInstantiate.Instantiate inName:=inName, inSource:=inSource, inDest:=inDest
End Function

Public Function generatorFromInstantiate(ByVal inProfile As FlnProfile, ByVal inAssets As FlnAssetList, ByVal inDebug As Boolean) As FlnGenerator
    Set generatorFromInstantiate = New FlnGenerator
    generatorFromInstantiate.Instantiate inProfile:=inProfile, inAssets:=inAssets, inDebug:=inDebug
End Function

Public Function outputTableFromInstantiate(ByVal inName As String, ByVal inIdList As Range, ByVal inSource As Range, ByVal inDest As Range) As FlnOutputTable
    Set outputTableFromInstantiate = New FlnOutputTable
    outputTableFromInstantiate.Instantiate inName:=inName, inIdList:=inIdList, inSource:=inSource, inDest:=inDest
End Function

Public Function pageGroupFromInstantiate(ByVal inName As String, ByVal inRange As String) As FlnPageGroup
    Set pageGroupFromInstantiate = New FlnPageGroup
    pageGroupFromInstantiate.Instantiate inName:=inName, inRange:=inRange
End Function

Public Function photoFillFromInstantiate(ByVal inName As String, ByVal inSource As Range, ByVal inDest As Range) As FlnPhotoFill
    Set photoFillFromInstantiate = New FlnPhotoFill
    photoFillFromInstantiate.Instantiate inName:=inName, inSource:=inSource, inDest:=inDest
End Function

Public Function profileFromInstantiate(ByVal inName As String, ByVal inRules As Collection, ByVal inAssetLists As Collection, ByVal inOutputTables As Collection, ByVal inSwaps As Collection, ByVal inPhotoFills As Collection, ByVal inAutofits As Collection, ByVal inAutohides As Collection, ByVal inPageGroups As Collection, ByVal inAssetCell As Range, ByVal inFilenameCell As Range, ByVal inPhotoPath As String, ByVal inPhotoRename As Range) As FlnProfile
    Set profileFromInstantiate = New FlnProfile
    profileFromInstantiate.Instantiate inName:=inName, inRules:=inRules, inAssetLists:=inAssetLists, inOutputTables:=inOutputTables, inSwaps:=inSwaps, inPhotoFills:=inPhotoFills, inAutofits:=inAutofits, inAutohides:=inAutohides, inPageGroups:=inPageGroups, inAssetCell:=inAssetCell, inFilenameCell:=inFilenameCell, inPhotoPath:=inPhotoPath, inPhotoRename:=inPhotoRename
End Function

Public Function ruleFromInstantiate(ByVal inName As String, ByVal inCases As Collection) As FlnRule
    Set ruleFromInstantiate = New FlnRule
    ruleFromInstantiate.Instantiate inName:=inName, inCases:=inCases
End Function

Public Function swapFromInstantiate(ByVal inName As String, ByVal inDoEndOrder As Boolean, ByVal inSwapCell As Range, ByVal inMaxSwapSet As Range, ByVal inSwapString As String, ByVal inSwapSet As Name, ByVal inSheets As Collection, ByVal inCValTables As Collection, ByVal inDoDupeSplit As Collection) As FlnSwap
    Set swapFromInstantiate = New FlnSwap
    swapFromInstantiate.Instantiate inName:=inName, inDoEndOrder:=inDoEndOrder, inSwapCell:=inSwapCell, inMaxSwapSet:=inMaxSwapSet, inSwapString:=inSwapString, inSwapSet:=inSwapSet, inSheets:=inSheets, inCValTables:=inCValTables, inDoDupeSplit:=inDoDupeSplit
End Function
