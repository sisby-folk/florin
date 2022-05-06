VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FlnUI 
   Caption         =   "Florin Report Generator"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   OleObjectBlob   =   "FlnUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FlnUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

' Colour Constants
Const UnlockedText As Long = 0
Const LockedText As Long = 7895160
Const InvalidText As Long = &HFF&

Private Profiles As Collection
Private CurrentProfile As FlnProfile

Private statusLines As Variant

''' Generator Tab Functions '''

Private Sub LockControl(ByVal ctrl As Object, ByVal newLocked As Boolean)
    ctrl.Enabled = Not newLocked
    ctrl.ForeColor = IIf(newLocked, LockedText, UnlockedText)
End Sub

Private Sub RunGenerator()
    If (LabelSaveWarning.Visible) Then           ' Validate Save
        MsgBox "Rules must be saved before run", vbOKOnly & vbExclamation, "Task Failed"
        ResetForm
        Exit Sub
    ElseIf ListGeneratorProfiles.ListIndex = -1 Or ListGeneratorAssetList.ListIndex = -1 Then
        MsgBox "Please select a Profile and Asset List to run", vbOKOnly & vbExclamation, "Task Failed"
        ResetForm
        Exit Sub
    Else
        Dim selectedProfile As FlnProfile
        Dim selectedAssetList As FlnAssetList
        Set selectedProfile = Profiles(ListGeneratorProfiles.ListIndex + 1)
        Set selectedAssetList = selectedProfile.AssetLists(ListGeneratorAssetList.ListIndex + 1)
        
        ' Check Basic Settings are set
        Dim unsetBasic As String
        unsetBasic = IIf(selectedProfile.AssetCell Is Nothing, "Asset Cell", vbNullString) & IIf(selectedProfile.FilenameCell Is Nothing And selectedProfile.GetSheets.count > 0, ", Filename Cell", vbNullString) & IIf(selectedProfile.photoPath = vbNullString And (selectedProfile.photofills.count > 0 Or Not selectedProfile.PhotoRename Is Nothing), ", Photo Path", vbNullString)
        If Not unsetBasic = vbNullString Then
            MsgBox unsetBasic & " not set. Please set these values in the Basic tab before continuing.", vbOKOnly & vbExclamation, "Task Failed"
            ResetForm
            Exit Sub
        End If
            
        ' Configure UI for script-running mode
        LockControl PageSwitcher, True
        LockControl GenerateButton, True
        LockControl ButtonCleanup, True
        Me.Height = Me.Height - FrameGeneratorTools.Height - 18
        statusLines = arrayFrom1DArgs(vbNullString, vbNullString, vbNullString, vbNullString)
        SetStatus Nothing, "Starting..."
        
        ' Initialize and Run Generator
        generatorFromInstantiate(selectedProfile, selectedAssetList, CheckDebug.value).Run
    End If
End Sub

Private Sub CleanupGenerator()
    ' Configure UI for script-running mode
    LockControl PageSwitcher, True
    LockControl GenerateButton, True
    LockControl ButtonCleanup, True
    Me.Height = Me.Height - FrameGeneratorTools.Height - 18
    statusLines = arrayFrom1DArgs(vbNullString, vbNullString, vbNullString, vbNullString)
    SetStatus Nothing, "Applied individual cleanup."
    
    ' Initialize and Run Generator
    If ListGeneratorProfiles.ListIndex = -1 Then
        generatorFromInstantiate(Nothing, Nothing, CheckDebug.value).CleanUp
    Else
        Dim selectedProfile As FlnProfile
        Set selectedProfile = Profiles(ListGeneratorProfiles.ListIndex + 1)
        generatorFromInstantiate(selectedProfile, Nothing, CheckDebug.value).CleanUp
    End If
End Sub

Public Sub ResetForm()
    ' Configure UI to default (ready to run) mode
    LockControl PageSwitcher, False
    LockControl GenerateButton, False
    LockControl ButtonCleanup, False
    PageSwitcher_Change
End Sub

Public Sub SetStatus(ByVal inFile As TextStream, Optional ByVal inLine1 As Variant, Optional ByVal inLine2 As Variant, Optional ByVal inLine3 As Variant, Optional ByVal inLine4 As Variant)
    ' Change the status label (can be called from generator)
    If Not IsMissing(inLine1) Then statusLines(1) = CStr(inLine1)
    If Not IsMissing(inLine2) Then statusLines(2) = CStr(inLine2)
    If Not IsMissing(inLine3) Then statusLines(3) = CStr(inLine3)
    If Not IsMissing(inLine4) Then statusLines(4) = CStr(inLine4)
    StatusLabel.Caption = arrayGetJoin(statusLines, vbNewLine)
    If Not inFile Is Nothing Then
        If Not IsMissing(inLine1) Then If inLine1 <> vbNullString Then inFile.WriteLine inLine1
        If Not IsMissing(inLine2) Then If inLine2 <> vbNullString Then inFile.WriteLine "   " & inLine2
        If Not IsMissing(inLine3) Then If inLine3 <> vbNullString Then inFile.WriteLine "       " & inLine3
        If Not IsMissing(inLine4) Then If inLine4 <> vbNullString Then inFile.WriteLine "           " & inLine4
    End If
End Sub

Private Sub ButtonOpenLogs_Click()
    Dim fso As FileSystemObject
    Dim logFolderPath As String
    Set fso = CreateObject("scripting.filesystemobject")
    logFolderPath = CreateObject("Wscript.Shell").SpecialFolders("MyDocuments") & "\FlorinLogs"
    If Not fso.FolderExists(logFolderPath) Then fso.CreateFolder logFolderPath
    SvnOffice.openInExplorer logFolderPath
End Sub

''' Generator Tab Events '''

Private Sub GenerateButton_Click()
    RunGenerator
End Sub

Private Sub ButtonCleanup_Click()
    CleanupGenerator
End Sub

''' Settings Tab Functions '''

Private Sub DirtyRules()
    LabelSaveWarning.Visible = True
    loadSettings
End Sub

Private Sub CleanRules()
    LabelSaveWarning.Visible = False
    loadSettings
End Sub

Private Sub reOrder(ByVal changeBox As MSForms.ListBox, ByVal changeCol As Collection, ByVal changeIndex As Long, ByVal directionUp As Boolean)
    If collectionMoveDirection(changeCol, changeIndex, directionUp) Then
        changeBox.ListIndex = changeIndex - 1 + IIf(directionUp, -1, 1)
        loadSettings
        DirtyRules
    End If
End Sub

Private Sub loadProfiles(ByVal workbookProperties As Collection, ByVal worksheetProperties As Collection, ByVal namedRanges As Collection, ByVal outProfiles As Collection)
    On Error GoTo Handler
    Dim numProfiles As Long
    numProfiles = 1
    
    Dim Rules As Collection
    Dim Cases As Collection
    Dim AssetLists As Collection
    Dim OutputTables As Collection
    Dim Swaps As Collection
    Dim photofills As Collection
    Dim autofits As Collection
    Dim autohides As Collection
    Dim pageGroups As Collection
    Dim AssetCell As Range
    Dim PhotoRename As Range
    Dim photoPath As String
    Dim FilenameCell As Range

    ''''' Profiles '''''
    Do
        Dim profileName As Variant
        profileName = collectionGetPropertyValueOrEmpty(workbookProperties, "Profile" & CStr(numProfiles))
        If IsEmpty(profileName) Then Exit Do     ' No More
        
        Set Rules = New Collection
        Set AssetLists = New Collection
        Set OutputTables = New Collection
        Set Swaps = New Collection
        Set photofills = New Collection
        Set autofits = New Collection
        Set autohides = New Collection
        Set pageGroups = New Collection
        
        ''''' Rules '''''
        Dim currentRuleNum As Long
        Dim ruleName As Variant
        currentRuleNum = 1
        Do
            ' Rule Name
            ruleName = collectionGetPropertyValueOrEmpty(workbookProperties, "Rule_" & profileName & "_Num" & CStr(currentRuleNum))
            If IsEmpty(ruleName) Then Exit Do    ' No More
            
            ' Cases
            Set Cases = New Collection
            Dim currentCaseNum As Long
            currentCaseNum = 1
            Do
                ' Rule Condition
                Dim Condition As Variant
                Condition = collectionGetPropertyValueOrEmpty(namedRanges, "Condition_" & profileName & "_" & ruleName & "_Case" & CStr(currentCaseNum))
                If IsEmpty(Condition) Then Exit Do ' No More

                ' Show Sheets
                Dim showSheets As Collection
                Set showSheets = New Collection
                Dim currentProperty As SvnProperty
                For Each currentProperty In worksheetProperties
                    If currentProperty.Name = "Visible_" & profileName & "_" & ruleName & "_Case" & CStr(currentCaseNum) Then
                        showSheets.Add currentProperty.sheet, currentProperty.sheet.Name
                    End If
                Next
                
                ' Add Case
                Cases.Add caseFromInstantiate(inCondition:=Condition, _
                                              inHideRef:=multiRangeFromInstantiate(ThisWorkbook, collectionGetPropertyValueOrEmpty(namedRanges, "Hide_" & profileName & "_" & ruleName & "_Case" & CStr(currentCaseNum)), inUnionize:=True, inEntireRows:=True), _
                                              inSheets:=showSheets)
                currentCaseNum = currentCaseNum + 1
            Loop
            
            ' Add Rule
            Rules.Add ruleFromInstantiate(ruleName, Cases), ruleName
            currentRuleNum = currentRuleNum + 1
        Loop
        
        ''''' Asset Lists '''''
        Dim currentAssetListNum As Long
        Dim listName As Variant
        currentAssetListNum = 1
        Do
            listName = collectionGetPropertyValueOrEmpty(workbookProperties, "Assets_" & profileName & "_Num" & CStr(currentAssetListNum))
            If IsEmpty(listName) Then Exit Do    ' No More
            
            AssetLists.Add assetListFromInstantiate(inName:=listName, inRange:=collectionGetPropertyValueOrEmpty(namedRanges, "Assets_" & profileName & "_" & listName)), listName
            currentAssetListNum = currentAssetListNum + 1
        Loop
        
        ''''' Output Tables '''''
        Dim currentOutTableNum As Long
        Dim outTableName As Variant
        currentOutTableNum = 1
        Do
            outTableName = collectionGetPropertyValueOrEmpty(workbookProperties, "Output_" & profileName & "_Num" & CStr(currentOutTableNum))
            If IsEmpty(outTableName) Then Exit Do ' No More

            OutputTables.Add outputTableFromInstantiate(inName:=outTableName, _
                                                        inIdList:=rangeOrNothing(collectionGetPropertyValueOrEmpty(namedRanges, "Output_" & profileName & "_" & outTableName & "_IDList"), ThisWorkbook), _
                                                        inSource:=rangeOrNothing(collectionGetPropertyValueOrEmpty(namedRanges, "Output_" & profileName & "_" & outTableName & "_Source"), ThisWorkbook), _
                                                        inDest:=rangeOrNothing(collectionGetPropertyValueOrEmpty(namedRanges, "Output_" & profileName & "_" & outTableName & "_Dest"), ThisWorkbook)), outTableName
            currentOutTableNum = currentOutTableNum + 1
        Loop
        
        
        
        ''''' Asset Cell '''''
        Dim AssetAddress As Variant
        AssetAddress = collectionGetPropertyValueOrEmpty(namedRanges, "AssetCell_" & profileName)
        Set AssetCell = rangeOrNothing(AssetAddress, ThisWorkbook)
        
        ''''' Photo Rename Range '''''
        Dim PhotoRenameAddress As Variant
        PhotoRenameAddress = collectionGetPropertyValueOrEmpty(namedRanges, "PhotoRename_" & profileName)
        Set PhotoRename = rangeOrNothing(PhotoRenameAddress, ThisWorkbook)
        
        ''''' Photo Path '''''
        Dim PhotoPathVariant As Variant
        PhotoPathVariant = collectionGetPropertyValueOrEmpty(workbookProperties, "PhotoPath_" & profileName)
        If IsEmpty(PhotoPathVariant) Then PhotoPathVariant = vbNullString
        photoPath = PhotoPathVariant
        
        ''''' Filename Cell '''''
        Dim FilenameAddress As Variant
        FilenameAddress = collectionGetPropertyValueOrEmpty(namedRanges, "FilenameCell_" & profileName)
        Set FilenameCell = rangeOrNothing(FilenameAddress, ThisWorkbook)
        
        ''''' Swaps '''''
        Dim currentSwapNum As Long
        Dim swapName As Variant
        currentSwapNum = 1
        Do
            swapName = collectionGetPropertyValueOrEmpty(workbookProperties, "Swap_" & profileName & "_Num" & CStr(currentSwapNum))
            If IsEmpty(swapName) Then Exit Do    ' No More
            
            ' Dupes
            Dim DupeSheets As Collection
            Set DupeSheets = New Collection
            Dim DupeDoSplits As Collection
            Set DupeDoSplits = New Collection
            For Each currentProperty In worksheetProperties
                If currentProperty.Name = "Swap_" & profileName & "_" & swapName & "_Dupe" Then
                    DupeSheets.Add currentProperty.sheet, currentProperty.sheet.Name
                ElseIf currentProperty.Name = "Swap_" & profileName & "_" & swapName & "_DupeSplit" Then
                    DupeDoSplits.Add CBool(currentProperty.value), currentProperty.sheet.Name
                End If
            Next
            If DupeSheets.count > 0 And DupeDoSplits.count = 0 Then Set DupeDoSplits = collectionFromKeys(sheetsToNames(DupeSheets), False)
            
            ' Cval Tables
            Dim CValTables As Collection
            Set CValTables = New Collection
            Dim currentTableNum As Long
            Dim tableName As Variant
            currentTableNum = 1
            Do
                tableName = collectionGetPropertyValueOrEmpty(workbookProperties, "Swap_" & profileName & "_" & swapName & "_CVal" & CStr(currentTableNum))
                If IsEmpty(tableName) Then Exit Do ' No More
                
                CValTables.Add cValTableFromInstantiate(inName:=tableName, _
                                                        inSource:=rangeOrNothing(collectionGetPropertyValueOrEmpty(namedRanges, "Swap_" & profileName & "_" & swapName & "_CVal_" & tableName & "_Source"), ThisWorkbook), _
                                                        inDest:=rangeOrNothing(collectionGetPropertyValueOrEmpty(namedRanges, "Swap_" & profileName & "_" & swapName & "_CVal_" & tableName & "_Dest"), ThisWorkbook)), tableName
                currentTableNum = currentTableNum + 1
            Loop
            
            ' Add Swap
            Swaps.Add swapFromInstantiate(inName:=swapName, _
                                          inDoEndOrder:=collectionGetPropertyValueOrEmpty(workbookProperties, "Swap_" & profileName & "_" & swapName & "_EndOrder"), _
                                          inSwapCell:=rangeOrNothing(collectionGetPropertyValueOrEmpty(namedRanges, "Swap_" & profileName & "_" & swapName & "_SwapCell"), ThisWorkbook), _
                                          inMaxSwapSet:=rangeOrNothing(collectionGetPropertyValueOrEmpty(namedRanges, "Swap_" & profileName & "_" & swapName & "_MaxSwapSet"), ThisWorkbook), _
                                          inSwapString:=collectionGetPropertyValueOrEmpty(namedRanges, "Swap_" & profileName & "_" & swapName & "_SwapSet"), _
                                          inSwapSet:=ThisWorkbook.Names("Rules_Swap_" & profileName & "_" & swapName & "_SwapSet"), _
                                          inSheets:=DupeSheets, _
                                          inCValTables:=CValTables, _
                                          inDoDupeSplit:=DupeDoSplits), swapName
            currentSwapNum = currentSwapNum + 1
        Loop
        
        ''''' PhotoFills '''''
        Dim currentPhotoNum As Long
        Dim photoName As Variant
        currentPhotoNum = 1
        Do
            photoName = collectionGetPropertyValueOrEmpty(workbookProperties, "Photo_" & profileName & "_Num" & CStr(currentPhotoNum))
            If IsEmpty(photoName) Then Exit Do   ' No More
            
            photofills.Add photoFillFromInstantiate(inName:=photoName, _
                                                    inSource:=rangeOrNothing(collectionGetPropertyValueOrEmpty(namedRanges, "Photo_" & profileName & "_" & photoName & "_Source"), ThisWorkbook), _
                                                    inDest:=rangeOrNothing(collectionGetPropertyValueOrEmpty(namedRanges, "Photo_" & profileName & "_" & photoName & "_Dest"), ThisWorkbook)), photoName
            currentPhotoNum = currentPhotoNum + 1
        Loop
        
        ''''' AutoFits '''''
        Dim currentAutofitNum As Long
        Dim autoFitName As Variant
        currentAutofitNum = 1
        Do
            autoFitName = collectionGetPropertyValueOrEmpty(workbookProperties, "Autofit_" & profileName & "_Num" & CStr(currentAutofitNum))
            If IsEmpty(autoFitName) Then Exit Do ' No More

            autofits.Add autoFitFromInstantiate(inName:=autoFitName, _
                                                inRange:=collectionGetPropertyValueOrEmpty(namedRanges, "Autofit_" & profileName & "_" & autoFitName & "_Range")), autoFitName
            currentAutofitNum = currentAutofitNum + 1
        Loop
        
        ''''' AutoHides '''''
        Dim currentAutohideNum As Long
        Dim autoHideName As Variant
        currentAutohideNum = 1
        Do
            autoHideName = collectionGetPropertyValueOrEmpty(workbookProperties, "Autohide_" & profileName & "_Num" & CStr(currentAutohideNum))
            If IsEmpty(autoHideName) Then Exit Do ' No More
            
            autohides.Add autoHideFromInstantiate(inName:=autoHideName, _
                                                  inRange:=collectionGetPropertyValueOrEmpty(namedRanges, "Autohide_" & profileName & "_" & autoHideName & "_Range")), autoHideName
            currentAutohideNum = currentAutohideNum + 1
        Loop
        
        ''''' PageGroups '''''
        Dim currentPageGroupNum As Long
        currentPageGroupNum = 1
        Dim pageGroupName As Variant
        Do
            pageGroupName = collectionGetPropertyValueOrEmpty(workbookProperties, "PageGroup_" & profileName & "_Num" & CStr(currentPageGroupNum))
            If IsEmpty(pageGroupName) Then Exit Do ' No More
            
            pageGroups.Add pageGroupFromInstantiate(pageGroupName, collectionGetPropertyValueOrEmpty(namedRanges, "PageGroup_" & profileName & "_" & pageGroupName & "_Range")), pageGroupName
            currentPageGroupNum = currentPageGroupNum + 1
        Loop
        
        ''''' Add Profile '''''
        outProfiles.Add profileFromInstantiate(profileName, Rules, AssetLists, OutputTables, Swaps, photofills, autofits, autohides, pageGroups, AssetCell, FilenameCell, photoPath, PhotoRename), profileName
        numProfiles = numProfiles + 1
    Loop
    Exit Sub
Handler:
    If HandleError(Err) = vbYes Then
        Stop
        Resume
    Else
        ' Clean Up
        Unload Me
        End
    End If
End Sub

Private Sub LoadProperties()
    ' Clear old object collections
    Set Profiles = New Collection

    ' Load property collections
    Dim workbookProperties As Collection
    Dim worksheetProperties As Collection
    Dim namedRanges As Collection
    Set workbookProperties = workbookGetProperties(ThisWorkbook, "Rules_")
    Set worksheetProperties = workbookGetSheetProperties(ThisWorkbook, "Rules_")
    Set namedRanges = workbookGetNameProperties(ThisWorkbook, "Rules_")
    
    loadProfiles workbookProperties, worksheetProperties, namedRanges, Profiles
    
End Sub

Public Function GetProfilesPublic() As Collection
    ' Clear old object collections
    Set GetProfilesPublic = New Collection

    ' Load property collections
    Dim workbookProperties As Collection
    Dim worksheetProperties As Collection
    Dim namedRanges As Collection
    Set workbookProperties = workbookGetProperties(ThisWorkbook, "Rules_")
    Set worksheetProperties = workbookGetSheetProperties(ThisWorkbook, "Rules_")
    Set namedRanges = workbookGetNameProperties(ThisWorkbook, "Rules_")
    
    loadProfiles workbookProperties, worksheetProperties, namedRanges, GetProfilesPublic
End Function

Private Sub SaveProperties()
    Dim workbookProperties As Collection
    Dim worksheetProperties As Collection
    Dim namedRanges As Collection
    Set workbookProperties = New Collection
    Set worksheetProperties = New Collection
    Set namedRanges = New Collection
    
    Dim currentProfileNum As Long
    For currentProfileNum = 1 To Profiles.count
        Dim EachProfile As FlnProfile
        Set EachProfile = Profiles(currentProfileNum)
        
        ' Profile Name
        addPropertyWithKey workbookProperties, "Profile" & CStr(currentProfileNum), EachProfile.Name, msoPropertyTypeString
        
        ''''' Rules '''''
        Dim currentRuleNum As Long
        For currentRuleNum = 1 To EachProfile.Rules.count
            Dim currentRule As FlnRule
            Set currentRule = EachProfile.Rules(currentRuleNum)
            
            ' Rule Name
            addPropertyWithKey workbookProperties, "Rule_" & EachProfile.Name & "_Num" & CStr(currentRuleNum), currentRule.Name, msoPropertyTypeString
            
            ' Cases
            Dim currentCaseNum As Long
            For currentCaseNum = 1 To currentRule.Cases.count
                Dim currentCase As FlnCase
                Set currentCase = currentRule.Cases(currentCaseNum)
                
                ' Case Condition
                If Not IsEmpty(currentCase.Condition) Then addPropertyWithKey namedRanges, "Condition_" & EachProfile.Name & "_" & currentRule.Name & "_Case" & CStr(currentCaseNum), currentCase.Condition, msoPropertyTypeString
                
                ' Hide Ranges
                If currentCase.HideRef.ToString() <> vbNullString Then addPropertyWithKey namedRanges, "Hide_" & EachProfile.Name & "_" & currentRule.Name & "_Case" & CStr(currentCaseNum), currentCase.HideRef.ToString, msoPropertyTypeString, inRange:=currentCase.HideRef
                
                ' Show Sheets
                Dim currentSheet As Worksheet
                For Each currentSheet In currentCase.Sheets
                    worksheetProperties.Add propertyFromInstantiate("Visible_" & EachProfile.Name & "_" & currentRule.Name & "_Case" & CStr(currentCaseNum), True, msoPropertyTypeString, currentSheet)
                Next
            Next
        Next
        
        ''''' Asset Lists '''''
        Dim currentAssetListNum As Long
        For currentAssetListNum = 1 To EachProfile.AssetLists.count
            Dim CurrentAssetList As FlnAssetList
            Set CurrentAssetList = EachProfile.AssetLists(currentAssetListNum)
            
            addPropertyWithKey workbookProperties, "Assets_" & EachProfile.Name & "_Num" & CStr(currentAssetListNum), CurrentAssetList.Name, msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Assets_" & EachProfile.Name & "_" & CurrentAssetList.Name, CurrentAssetList.ToString, msoPropertyTypeString, inRange:=CurrentAssetList.MultiRange
        Next
        
        ' Output Tables
        Dim currentOutputTableNum As Long
        For currentOutputTableNum = 1 To EachProfile.OutputTables.count
            Dim currentOutputTable As FlnOutputTable
            Set currentOutputTable = EachProfile.OutputTables(currentOutputTableNum)
            
            addPropertyWithKey workbookProperties, "Output_" & EachProfile.Name & "_Num" & CStr(currentOutputTableNum), currentOutputTable.Name, msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Output_" & EachProfile.Name & "_" & currentOutputTable.Name & "_IDList", rangeGetAddress(currentOutputTable.IdList), msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Output_" & EachProfile.Name & "_" & currentOutputTable.Name & "_Source", rangeGetAddress(currentOutputTable.Source), msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Output_" & EachProfile.Name & "_" & currentOutputTable.Name & "_Dest", rangeGetAddress(currentOutputTable.Dest), msoPropertyTypeString
        Next
        
        ''''' Asset Cell '''''
        If (Not EachProfile.AssetCell Is Nothing) Then addPropertyWithKey namedRanges, "AssetCell_" & EachProfile.Name, rangeGetAddress(EachProfile.AssetCell), msoPropertyTypeString
        
        ''''' Photo Rename Range '''''
        If (Not EachProfile.PhotoRename Is Nothing) Then addPropertyWithKey namedRanges, "PhotoRename_" & EachProfile.Name, rangeGetAddress(EachProfile.PhotoRename), msoPropertyTypeString
        
        ''''' Photo Path '''''
        addPropertyWithKey workbookProperties, "PhotoPath_" & EachProfile.Name, EachProfile.photoPath, msoPropertyTypeString
        
        ''''' Filename Cell '''''
        If (Not EachProfile.FilenameCell Is Nothing) Then addPropertyWithKey namedRanges, "FilenameCell_" & EachProfile.Name, rangeGetAddress(EachProfile.FilenameCell), msoPropertyTypeString
        
        ''''' Swaps '''''
        Dim currentSwapNum As Long
        Dim currentDupe As Worksheet
        For currentSwapNum = 1 To EachProfile.Swaps.count
            Dim currentSwap As FlnSwap
            Set currentSwap = EachProfile.Swaps(currentSwapNum)
            
            addPropertyWithKey workbookProperties, "Swap_" & EachProfile.Name & "_Num" & CStr(currentSwapNum), currentSwap.Name, msoPropertyTypeString
            
            addPropertyWithKey workbookProperties, "Swap_" & EachProfile.Name & "_" & currentSwap.Name & "_EndOrder", currentSwap.doEndOrder, msoPropertyTypeBoolean
            
            addPropertyWithKey namedRanges, "Swap_" & EachProfile.Name & "_" & currentSwap.Name & "_SwapCell", rangeGetAddress(currentSwap.SwapCell), msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Swap_" & EachProfile.Name & "_" & currentSwap.Name & "_SwapSet", currentSwap.swapString, msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Swap_" & EachProfile.Name & "_" & currentSwap.Name & "_MaxSwapSet", rangeGetAddress(currentSwap.MaxSwapSet), msoPropertyTypeString
            
            For Each currentDupe In currentSwap.DupeSheets
                worksheetProperties.Add propertyFromInstantiate("Swap_" & EachProfile.Name & "_" & currentSwap.Name & "_Dupe", True, msoPropertyTypeString, currentDupe)
                worksheetProperties.Add propertyFromInstantiate("Swap_" & EachProfile.Name & "_" & currentSwap.Name & "_DupeSplit", currentSwap.DoDupeSplit(currentDupe.Name), msoPropertyTypeString, currentDupe)
            Next
            
            ' CVal Tables
            Dim currentCValTableNum As Long
            For currentCValTableNum = 1 To currentSwap.CValTables.count
                Dim currentCValTable As FlnCValTable
                Set currentCValTable = currentSwap.CValTables(currentCValTableNum)
                
                addPropertyWithKey workbookProperties, "Swap_" & EachProfile.Name & "_" & currentSwap.Name & "_CVal" & CStr(currentCValTableNum), currentCValTable.Name, msoPropertyTypeString
                
                addPropertyWithKey namedRanges, "Swap_" & EachProfile.Name & "_" & currentSwap.Name & "_CVal_" & currentCValTable.Name & "_Source", rangeGetAddress(currentCValTable.Source), msoPropertyTypeString
                
                addPropertyWithKey namedRanges, "Swap_" & EachProfile.Name & "_" & currentSwap.Name & "_CVal_" & currentCValTable.Name & "_Dest", rangeGetAddress(currentCValTable.Dest), msoPropertyTypeString
            Next
        Next
        
        ''''' Photo Fills '''''
        Dim currentPhotoNum As Long
        For currentPhotoNum = 1 To EachProfile.photofills.count
            Dim currentPhotoFill As FlnPhotoFill
            Set currentPhotoFill = EachProfile.photofills(currentPhotoNum)
            
            addPropertyWithKey workbookProperties, "Photo_" & EachProfile.Name & "_Num" & CStr(currentPhotoNum), currentPhotoFill.Name, msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Photo_" & EachProfile.Name & "_" & currentPhotoFill.Name & "_Source", rangeGetAddress(currentPhotoFill.Source), msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Photo_" & EachProfile.Name & "_" & currentPhotoFill.Name & "_Dest", rangeGetAddress(currentPhotoFill.Dest), msoPropertyTypeString
        Next
        
        ''''' Autofit '''''
        Dim currentAutofitNum As Long
        For currentAutofitNum = 1 To EachProfile.autofits.count
            Dim currentAutofit As FlnAutoFit
            Set currentAutofit = EachProfile.autofits(currentAutofitNum)
            
            addPropertyWithKey workbookProperties, "Autofit_" & EachProfile.Name & "_Num" & CStr(currentAutofitNum), currentAutofit.Name, msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Autofit_" & EachProfile.Name & "_" & currentAutofit.Name & "_Range", currentAutofit.MultiRange.ToString, msoPropertyTypeString, inRange:=currentAutofit.MultiRange
        Next
         
        ''''' Autohide '''''
         
        Dim currentAutohideNum As Long
        For currentAutohideNum = 1 To EachProfile.autohides.count
            Dim currentAutohide As FlnAutoHide
            Set currentAutohide = EachProfile.autohides(currentAutohideNum)
            
            addPropertyWithKey workbookProperties, "Autohide_" & EachProfile.Name & "_Num" & CStr(currentAutohideNum), currentAutohide.Name, msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "Autohide_" & EachProfile.Name & "_" & currentAutohide.Name & "_Range", currentAutohide.MultiRange.ToString, msoPropertyTypeString, inRange:=currentAutohide.MultiRange
        Next
        
        ''''' PageGroup '''''
        Dim currentPageGroupNum As Long
        For currentPageGroupNum = 1 To EachProfile.pageGroups.count
            Dim currentPageGroup As FlnPageGroup
            Set currentPageGroup = EachProfile.pageGroups(currentPageGroupNum)
            
            addPropertyWithKey workbookProperties, "PageGroup_" & EachProfile.Name & "_Num" & CStr(currentPageGroupNum), currentPageGroup.Name, msoPropertyTypeString
            
            addPropertyWithKey namedRanges, "PageGroup_" & EachProfile.Name & "_" & currentPageGroup.Name & "_Range", currentPageGroup.MultiRange.ToString, msoPropertyTypeString, inRange:=currentPageGroup.MultiRange
        Next
    Next
    
    workbookOverwriteProperties ThisWorkbook, workbookProperties, "Rules_"
    workbookOverwriteSheetProperties ThisWorkbook, worksheetProperties, "Rules_"
    workbookOverwriteNamedRanges ThisWorkbook, namedRanges, "Rules_"
    LoadProperties
    CleanRules
End Sub

Private Sub EnableProfileEditing(ByVal NewEnabled As Boolean)
    LockControl FrameRules, Not NewEnabled
    LockControl FrameAssetList, Not NewEnabled
    LockControl FrameOutputs, Not NewEnabled
    LockControl FrameSwaps, Not NewEnabled
    LockControl FramePhotoFills, Not NewEnabled
    LockControl FrameAutofits, Not NewEnabled
    LockControl FrameAutohides, Not NewEnabled
    LockControl FramePageGroups, Not NewEnabled
    LockControl ButtonSetFilenameCell, Not NewEnabled
    LockControl ButtonClearFilenameCell, Not NewEnabled
    LockControl ButtonSetAssetCell, Not NewEnabled
    LockControl ButtonClearAssetCell, Not NewEnabled
    LockControl ButtonSetPhotoPath, Not NewEnabled
    LockControl ButtonClearPhotoPath, Not NewEnabled
    LockControl ButtonSetPhotoRename, Not NewEnabled
    LockControl ButtonClearPhotoRename, Not NewEnabled
    LockControl ToggleDupeSplit, Not NewEnabled
    LockControl ToggleSwapEnd, Not NewEnabled
    
    
    LockControl ButtonNewAutofit, Not NewEnabled
    LockControl ButtonDeleteAutofit, Not NewEnabled
    LockControl ButtonEditAutofit, Not NewEnabled
    
    LockControl ButtonNewAutohide, Not NewEnabled
    LockControl ButtonDeleteAutohide, Not NewEnabled
    LockControl ButtonEditAutohide, Not NewEnabled
    
    LockControl ButtonNewPageGroup, Not NewEnabled
    LockControl ButtonRemovePageGroup, Not NewEnabled
    LockControl ButtonEditPageGroup, Not NewEnabled
    
    LockControl ButtonNewCase, Not NewEnabled
    LockControl ButtonRemoveCase, Not NewEnabled
    LockControl ButtonEditCase, Not NewEnabled
    LockControl ButtonCaseUp, Not NewEnabled
    LockControl ButtonCaseDown, Not NewEnabled
    
    LockControl ButtonNewPhotoFill, Not NewEnabled
    LockControl ButtonRemovePhotoFill, Not NewEnabled
    LockControl ButtonEditPhotoFill, Not NewEnabled
    
    LockControl ButtonNewCVal, Not NewEnabled
    LockControl ButtonRemoveCVal, Not NewEnabled
    LockControl ButtonEditCValOutput, Not NewEnabled
    
    LockControl ButtonNewOutput, Not NewEnabled
    LockControl ButtonRemoveOutput, Not NewEnabled
    LockControl ButtonEditOutput, Not NewEnabled
    
    LockControl ButtonNewAssetList, Not NewEnabled
    LockControl ButtonRemoveAssetList, Not NewEnabled
    LockControl ButtonEditAssetList, Not NewEnabled
    
    LockControl ButtonNewRule, Not NewEnabled
    LockControl ButtonRemoveRule, Not NewEnabled
    LockControl ButtonRuleUp, Not NewEnabled
    LockControl ButtonRuleDown, Not NewEnabled
    LockControl ButtonDupeRule, Not NewEnabled
    LockControl ButtonEditRule, Not NewEnabled
    
    LockControl ButtonNewSwap, Not NewEnabled
    LockControl ButtonRemoveSwap, Not NewEnabled
    LockControl ButtonEditSwap, Not NewEnabled
    
    LockControl ButtonNewDupe, Not NewEnabled
    LockControl ButtonRemoveDupe, Not NewEnabled
End Sub

Private Sub reSelectGeneratorProfile()
    ListGeneratorAssetList.Clear
    If ListGeneratorProfiles.ListIndex <> -1 Then
        Dim selectedProfile As FlnProfile
        Set selectedProfile = Profiles(ListGeneratorProfiles.ListIndex + 1)
        Dim CurrentAssetList As FlnAssetList
        For Each CurrentAssetList In selectedProfile.AssetLists
            ListGeneratorAssetList.AddItem CurrentAssetList.Name
        Next
    End If
End Sub

Private Sub reLoadProfiles()                     ' Profiles have been loaded or edited
    ' Load Profiles - Combo & Generator
    ComboProfiles.Clear
    ListGeneratorProfiles.Clear
    Dim prof As FlnProfile
    For Each prof In Profiles
        ComboProfiles.AddItem prof.Name
        ListGeneratorProfiles.AddItem prof.Name
    Next
    reSelectProfile
End Sub

Private Sub reSelectProfile()                    ' Profile has been changed (but not reloaded)
    ' Clear all selections and reload
    ListRules.ListIndex = -1
    ListCases.ListIndex = -1
    ListAssetLists.ListIndex = -1
    ListOutputs.ListIndex = -1
    ListSwaps.ListIndex = -1
    loadSettings                                 ' Settings will use new profile index
End Sub

Private Sub reSelectAutofit()
    LockControl ButtonDeleteAutofit, ListAutofits.ListIndex = -1
    LockControl ButtonEditAutofit, ListAutofits.ListIndex = -1
End Sub

Private Sub reSelectAutohide()
    LockControl ButtonDeleteAutohide, ListAutohides.ListIndex = -1
    LockControl ButtonEditAutohide, ListAutohides.ListIndex = -1
End Sub

Private Sub reSelectPageGroup()
    LockControl ButtonRemovePageGroup, ListPageGroups.ListIndex = -1
    LockControl ButtonEditPageGroup, ListPageGroups.ListIndex = -1
End Sub

Private Sub reSelectCase()
    LockControl ButtonRemoveCase, ListCases.ListIndex = -1
    LockControl ButtonEditCase, ListCases.ListIndex = -1
    LockControl ButtonCaseUp, ListCases.ListIndex = -1
    LockControl ButtonCaseDown, ListCases.ListIndex = -1
End Sub

Private Sub reSelectPhotoFill()
    LockControl ButtonRemovePhotoFill, ListPhotoFills.ListIndex = -1
    LockControl ButtonEditPhotoFill, ListPhotoFills.ListIndex = -1
End Sub

Private Sub reSelectCVal()
    LockControl ButtonRemoveCVal, ListCVals.ListIndex = -1
    LockControl ButtonEditCValOutput, ListCVals.ListIndex = -1
End Sub

Private Sub reSelectOutput()
    LockControl ButtonRemoveOutput, ListOutputs.ListIndex = -1
    LockControl ButtonEditOutput, ListOutputs.ListIndex = -1
End Sub

Private Sub reSelectAssetList()
    LockControl ButtonRemoveAssetList, ListAssetLists.ListIndex = -1
    LockControl ButtonEditAssetList, ListAssetLists.ListIndex = -1
End Sub

Private Sub reSelectRule()
    ' Adjust Control Locks
    LockControl ButtonRemoveRule, ListRules.ListIndex = -1
    LockControl ButtonRuleUp, ListRules.ListIndex = -1
    LockControl ButtonRuleDown, ListRules.ListIndex = -1
    LockControl ButtonDupeRule, ListRules.ListIndex = -1
    LockControl ButtonEditRule, ListRules.ListIndex = -1
    LockControl ButtonNewCase, ListRules.ListIndex = -1
    
    ' Repopulate Cases
    ListCases.Clear
    If ListRules.ListIndex <> -1 Then
        Dim selectedRule As FlnRule
        Set selectedRule = CurrentProfile.Rules(ListRules.ListIndex + 1)
        LabelCases.Caption = "Cases: " & selectedRule.Name
        Dim currentCase As FlnCase
        For Each currentCase In selectedRule.Cases
            ListCases.AddItem
            ListCases.list(ListCases.ListCount - 1, 0) = IIf(currentCase.Sheets.count = 0, vbNullString, "Shows " & currentCase.Sheets.count & " Sheets")
            ListCases.list(ListCases.ListCount - 1, 1) = IIf(currentCase.HideRef.Worksheets.count = 0, vbNullString, "Hides Rows")
            ListCases.list(ListCases.ListCount - 1, 2) = currentCase.Condition
        Next
    Else
        LabelCases.Caption = "Cases: No Rule Selected"
    End If
End Sub

Private Sub reSelectSwap()
    LockControl ButtonRemoveSwap, ListSwaps.ListIndex = -1
    LockControl ButtonEditSwap, ListSwaps.ListIndex = -1

    ListDupes.Clear
    ListCVals.Clear
    ToggleSwapEnd.Tag = "Locked"
    
    ToggleSwapEnd.value = False
    ToggleSwapEnd.Enabled = False
    If ListSwaps.ListIndex <> -1 Then
        Dim selectedSwap As FlnSwap
        Set selectedSwap = CurrentProfile.Swaps(ListSwaps.ListIndex + 1)
        LabelDupes.Caption = "Dupes: " & selectedSwap.Name
        ToggleSwapEnd.Enabled = True
        ToggleSwapEnd.value = selectedSwap.doEndOrder
        Dim currentSheet As Worksheet
        For Each currentSheet In selectedSwap.DupeSheets
            ListDupes.AddItem currentSheet.Name
        Next
        LabelCVals.Caption = "CVal Tables: " & selectedSwap.Name
        Dim currentCVal As FlnCValTable
        For Each currentCVal In selectedSwap.CValTables
            ListCVals.AddItem currentCVal.Name
        Next
    Else
        LabelCVals.Caption = "No Swap Selected"
        LabelCases.Caption = "No Swap Selected"
    End If
    
    ToggleSwapEnd.Tag = vbNullString
End Sub

Private Sub reSelectDupe()
    LockControl ButtonRemoveDupe, ListDupes.ListIndex = -1
    
    ToggleDupeSplit.Tag = "Locked"
    
    ToggleDupeSplit.value = False
    ToggleDupeSplit.Enabled = False
    If ListSwaps.ListIndex <> -1 And ListDupes.ListIndex <> -1 Then
        Dim selectedSwap As FlnSwap
        Set selectedSwap = CurrentProfile.Swaps(ListSwaps.ListIndex + 1)
        ToggleDupeSplit.Enabled = True
        ToggleDupeSplit.value = selectedSwap.DoDupeSplit(ListDupes.ListIndex + 1) ' Keyed by Index
    End If
    
    ToggleDupeSplit.Tag = vbNullString
End Sub

Private Sub loadSettings()

    ' Save and Clear All
    Dim ruleIndex As Long
    Dim caseIndex As Long
    Dim AssetListIndex As Long
    Dim OutputIndex As Long
    Dim swapIndex As Long
    Dim DupeIndex As Long
    Dim CValIndex As Long
    Dim PhotoFillIndex As Long
    Dim AutofitIndex As Long
    Dim AutohideIndex As Long
    Dim PageGroupIndex As Long
    
    ruleIndex = ListRules.ListIndex
    caseIndex = ListCases.ListIndex
    AssetListIndex = ListAssetLists.ListIndex
    swapIndex = ListSwaps.ListIndex
    DupeIndex = ListDupes.ListIndex
    OutputIndex = ListOutputs.ListIndex
    CValIndex = ListCVals.ListIndex
    PhotoFillIndex = ListPhotoFills.ListIndex
    AutofitIndex = ListAutofits.ListIndex
    AutohideIndex = ListAutohides.ListIndex
    PageGroupIndex = ListPageGroups.ListIndex
    
    ListRules.Clear
    ListAssetLists.Clear
    ListOutputs.Clear
    ListSwaps.Clear
    ListPhotoFills.Clear
    ListAutofits.Clear
    ListAutohides.Clear
    ListPageGroups.Clear
    LabelAssetCell.Caption = vbNullString
    LabelPhotoRename.Caption = vbNullString
    LabelPhotoPath.Caption = vbNullString
    LabelFilenameCell.Caption = vbNullString

    
    Set CurrentProfile = collectionGetOrNothing(Profiles, ComboProfiles.value) ' Reload currentProfile to remove dangling references
    
    ' Validate Profile Existence
    If Not CurrentProfile Is Nothing Then
        EnableProfileEditing True

        ' Load Rules
        Dim currentRule As FlnRule
        For Each currentRule In CurrentProfile.Rules
            ListRules.AddItem currentRule.Name
        Next
        If ruleIndex < ListRules.ListCount Then ListRules.ListIndex = ruleIndex
        reSelectRule
        If caseIndex < ListCases.ListCount Then ListCases.ListIndex = caseIndex
        reSelectCase
        
        ' Load Asset Lists
        Dim CurrentAssetList As FlnAssetList
        For Each CurrentAssetList In CurrentProfile.AssetLists
            ListAssetLists.AddItem
            ListAssetLists.list(ListAssetLists.ListCount - 1, 0) = CurrentAssetList.MultiRange.Areas(1).Rows.count
            ListAssetLists.list(ListAssetLists.ListCount - 1, 1) = CurrentAssetList.Name
            ListAssetLists.list(ListAssetLists.ListCount - 1, 2) = CurrentAssetList.ToString
        Next
        If AssetListIndex < ListAssetLists.ListCount Then ListAssetLists.ListIndex = AssetListIndex
        reSelectAssetList
        
        ' Load Output Tables
        Dim currentOutput As FlnOutputTable
        For Each currentOutput In CurrentProfile.OutputTables
            ListOutputs.AddItem
            ListOutputs.list(ListOutputs.ListCount - 1, 0) = currentOutput.Name
            ListOutputs.list(ListOutputs.ListCount - 1, 1) = currentOutput.Source.Worksheet.Name
            ListOutputs.list(ListOutputs.ListCount - 1, 2) = currentOutput.Source.Rows.count
            ListOutputs.list(ListOutputs.ListCount - 1, 3) = rangeGetAddress(currentOutput.IdList)
        Next
        If OutputIndex < ListOutputs.ListCount Then ListOutputs.ListIndex = OutputIndex
        reSelectOutput
        
        ' Load Swaps
        Dim currentSwap As FlnSwap
        For Each currentSwap In CurrentProfile.Swaps
            ListSwaps.AddItem
            ListSwaps.list(ListSwaps.ListCount - 1, 0) = currentSwap.Name
            ListSwaps.list(ListSwaps.ListCount - 1, 1) = IIf(currentSwap.doEndOrder, "End", "Before")
            ListSwaps.list(ListSwaps.ListCount - 1, 2) = rangeGetAddress(currentSwap.SwapCell)
            ListSwaps.list(ListSwaps.ListCount - 1, 3) = currentSwap.swapString
        Next
        If swapIndex < ListSwaps.ListCount Then ListSwaps.ListIndex = swapIndex
        reSelectSwap
        If DupeIndex < ListDupes.ListCount Then ListDupes.ListIndex = DupeIndex
        reSelectDupe
        If CValIndex < ListCVals.ListCount Then ListCVals.ListIndex = CValIndex
        reSelectCVal
        
        ' Load Photo Fills
        Dim currentPhotoFill As FlnPhotoFill
        For Each currentPhotoFill In CurrentProfile.photofills
            ListPhotoFills.AddItem
            ListPhotoFills.list(ListPhotoFills.ListCount - 1, 0) = currentPhotoFill.Name
            ListPhotoFills.list(ListPhotoFills.ListCount - 1, 1) = currentPhotoFill.Dest.Areas.count
            ListPhotoFills.list(ListPhotoFills.ListCount - 1, 2) = currentPhotoFill.Dest.Address
        Next
        If PhotoFillIndex < ListPhotoFills.ListCount Then ListPhotoFills.ListIndex = PhotoFillIndex
        reSelectPhotoFill
        
        
        ' Load Autofits
        Dim currentAutofit As FlnAutoFit
        For Each currentAutofit In CurrentProfile.autofits
            ListAutofits.AddItem
            ListAutofits.list(ListAutofits.ListCount - 1, 0) = currentAutofit.Name
            ListAutofits.list(ListAutofits.ListCount - 1, 1) = currentAutofit.MultiRange.Pretty
        Next
        If AutofitIndex < ListAutofits.ListCount Then ListAutofits.ListIndex = AutofitIndex
        reSelectAutofit
        
        ' Load Autohides
        Dim currentAutohide As FlnAutoHide
        For Each currentAutohide In CurrentProfile.autohides
            ListAutohides.AddItem
            ListAutohides.list(ListAutohides.ListCount - 1, 0) = currentAutohide.Name
            ListAutohides.list(ListAutohides.ListCount - 1, 1) = currentAutohide.MultiRange.Pretty
        Next
        If AutohideIndex < ListAutohides.ListCount Then ListAutohides.ListIndex = AutohideIndex
        reSelectAutohide
        
        ' Load PageGroups
        Dim currentPageGroup As FlnPageGroup
        For Each currentPageGroup In CurrentProfile.pageGroups
            ListPageGroups.AddItem
            ListPageGroups.list(ListPageGroups.ListCount - 1, 0) = currentPageGroup.Name
            ListPageGroups.list(ListPageGroups.ListCount - 1, 1) = currentPageGroup.MultiRange.Pretty
        Next
        If PageGroupIndex < ListPageGroups.ListCount Then ListPageGroups.ListIndex = PageGroupIndex
        reSelectPageGroup
        
        ' Load Basic Tab
        If CurrentProfile.AssetCell Is Nothing Then labelSetContent LabelAssetCell, "Not set", InvalidText Else labelSetContent LabelAssetCell, rangeGetAddress(CurrentProfile.AssetCell), UnlockedText
        
        If CurrentProfile.PhotoRename Is Nothing Then labelSetContent LabelPhotoRename, "Not set", InvalidText Else labelSetContent LabelPhotoRename, rangeGetAddress(CurrentProfile.PhotoRename), UnlockedText
        
        If CurrentProfile.FilenameCell Is Nothing Then labelSetContent LabelFilenameCell, "Not set", InvalidText Else labelSetContent LabelFilenameCell, rangeGetAddress(CurrentProfile.FilenameCell), UnlockedText
        
        If CurrentProfile.photoPath = vbNullString Then labelSetContent LabelPhotoPath, "Not set", InvalidText Else labelSetContent LabelPhotoPath, CurrentProfile.photoPath, UnlockedText
        
    Else
        EnableProfileEditing False
    End If
End Sub

Public Sub EditCase()
    If ListRules.ListIndex <> -1 And ListCases.ListIndex <> -1 Then
        Dim selectedRule As FlnRule
        Dim selectedCase As FlnCase
        Set selectedRule = CurrentProfile.Rules(ListRules.ListIndex + 1)
        Set selectedCase = selectedRule.Cases(ListCases.ListIndex + 1)
        With New FlnCaseEditor
            .Initialize caseFromInstantiate(selectedCase.Condition, selectedCase.HideRef, selectedCase.Sheets), _
        selectedRule.Name, ListCases.ListIndex
            Me.Hide
            .Show
            If .isDirty Then
                selectedCase.Instantiate .FormCase.Condition, .FormCase.HideRef, .FormCase.Sheets
                DirtyRules
            End If
        End With
        Me.Show
    End If
End Sub

Public Sub EditProfiles()
    With New FlnProfileEditor
        .Initialize Profiles
        Me.Hide
        .Show
        If .isDirty Then
            reLoadProfiles
            DirtyRules
        End If
        If .ReturnIndex <> -1 Then
            ComboProfiles.ListIndex = .ReturnIndex
            reSelectProfile
        End If
    End With
End Sub

Public Sub EditPhotoFill(ByVal selectedPhotoFill As FlnPhotoFill)
    Dim NewSourceString As String
    Dim NewSource As Range
    NewSourceString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(selectedPhotoFill.Source), "Edit Photo Fill Source Range", Dimensions:=dimensionsFromRowsColsRepeatable(False, 1, 1)) ' Can't be cross-sheet, must preserve order
    If NewSourceString = vbNullString Then GoTo Err
    
    Set NewSource = rangeOrNothing(NewSourceString, ThisWorkbook)
    
    Dim NewDestString As String
    NewDestString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(selectedPhotoFill.Dest), "Edit Photo Fill Dest Range", Dimensions:=dimensionsFromRowsColsAreas(-1, -1, NewSource.Areas.count))
    If NewDestString = vbNullString Then GoTo Err
    
    selectedPhotoFill.Instantiate selectedPhotoFill.Name, NewSource, rangeOrNothing(NewDestString, ThisWorkbook)
    DirtyRules
    Exit Sub
Err:
    MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
End Sub

Public Sub EditAssetList(ByVal selectedAssetList As FlnAssetList)
    Dim NewRangeString As String
    NewRangeString = stringFromRefEdit(ThisWorkbook, selectedAssetList.ToString, "Edit Asset List", Dimensions:=dimensionsFromRowsCols(-1, 1), Unionize:=True)
    If NewRangeString = vbNullString Then GoTo Err
    
    selectedAssetList.setRange NewRangeString
    DirtyRules
    Exit Sub
Err:
    MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
End Sub

Public Sub EditCValTable(ByVal selectedSwap As FlnSwap, ByVal selectedTable As FlnCValTable)
    Dim SourceString As String
    SourceString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(selectedTable.Source), "CVal Source - Swap 1", Dimensions:=dimensionsFromRowsCols(-1, -1))
    If SourceString = vbNullString Then GoTo Err
    Dim SourceRange As Range
    Set SourceRange = rangeOrNothing(SourceString, ThisWorkbook)
    
    Dim DestString As String
    DestString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(selectedTable.Dest), "CVal Dest - Areas for Swap 2 Onwards", Dimensions:=dimensionsFromRowsColsRepeatable(False, SourceRange.Rows.count, SourceRange.Columns.count))
    If DestString = vbNullString Then GoTo Err
    Dim DestRange As Range
    Set DestRange = rangeOrNothing(DestString, ThisWorkbook)
    If DestRange.Areas.count + 1 <> selectedSwap.MaxSwapSet.Rows.count Then GoTo Err ' Not enough space to copy

    selectedTable.Instantiate selectedTable.Name, SourceRange, DestRange
    DirtyRules
    Exit Sub
Err:
    MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
End Sub

Public Sub EditOutputTable(ByVal selectedTable As FlnOutputTable)
    Dim IDString As String
    IDString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(selectedTable.IdList), "Sorted List of IDs", Dimensions:=dimensionsFromRowsCols(-1, 1))
    If IDString = vbNullString Then GoTo Err
    
    Dim SourceString As String
    SourceString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(selectedTable.Source), "Output Source - Sigs, Data, and Valid", Dimensions:=dimensionsFromRowsCols(-1, -1, -1, -1, -1, 1))
    If SourceString = vbNullString Then GoTo Err
    Dim SourceRange As Range
    Set SourceRange = rangeOrNothing(SourceString, ThisWorkbook) ' Should work. Offensive.
    
    Dim DestString As String
    DestString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(selectedTable.Dest), "Output Dest - IDs, Sigs, and Data", Dimensions:=dimensionsFromRowsCols(-1, 1, -1, SourceRange.Areas(1).Columns.count, -1, SourceRange.Areas(2).Columns.count))
    If DestString = vbNullString Then GoTo Err
    
    selectedTable.Instantiate selectedTable.Name, rangeOrNothing(IDString, ThisWorkbook), SourceRange, rangeOrNothing(DestString, ThisWorkbook)
    DirtyRules
    Exit Sub
Err:
    MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
End Sub

Public Sub EditSwap(ByVal selectedSwap As FlnSwap)
    Dim CellString As String
    CellString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(selectedSwap.SwapCell), "Swap Cell", Dimensions:=dimensionsFromRowsCols(1, 1))
    If CellString = vbNullString Then GoTo Err
    
    Dim MaxString As String
    MaxString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(selectedSwap.MaxSwapSet), "Max SwapSet", Dimensions:=dimensionsFromRowsCols(1, -1))
    If MaxString = vbNullString Then GoTo Err
    
    Dim SetString As String
    SetString = stringFromFormulaEdit(selectedSwap.swapString, "Dynamic Swap Set")
    If SetString = vbNullString Then GoTo Err
    
    selectedSwap.Edit selectedSwap.doEndOrder, rangeOrNothing(CellString, ThisWorkbook), rangeOrNothing(MaxString, ThisWorkbook), SetString
    DirtyRules
    Exit Sub
Err:
    MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
End Sub

Public Sub EditAutofit(ByVal selectedAutofit As FlnAutoFit)
    If selectedAutofit.MultiRange.Edit("Autofit " & selectedAutofit.Name, dimensionsFromRowsColsRepeatable(True, -1, -1)) Then DirtyRules
End Sub

Public Sub EditAutohide(ByVal selectedAutohide As FlnAutoHide)
    If selectedAutohide.MultiRange.Edit("Autohide " & selectedAutohide.Name, dimensionsFromRowsColsRepeatable(True, -1, 1)) Then DirtyRules
End Sub

Public Sub EditPageGroup(ByVal selectedPageGroup As FlnPageGroup)
    If selectedPageGroup.MultiRange.Edit("PageGroup " & selectedPageGroup.Name, dimensionsFromRowsColsRepeatable(True, -1, -1)) Then DirtyRules
End Sub

''' Settings Tab Events '''

Private Sub ComboProfiles_Change()
    reSelectProfile
End Sub

Private Sub ListAssetLists_Click()
    reSelectAssetList
End Sub

Private Sub ListAutofits_Click()
    reSelectAutofit
End Sub

Private Sub ListAutohides_Click()
    reSelectAutohide
End Sub

Private Sub ListCases_Click()
    reSelectCase
End Sub

Private Sub ListCVals_Click()
    reSelectCVal
End Sub

Private Sub ListDupes_Click()
    reSelectDupe
End Sub

Private Sub ListGeneratorProfiles_Click()
    reSelectGeneratorProfile
End Sub

Private Sub ListOutputs_Click()
    reSelectOutput
End Sub

Private Sub ListPageGroups_Click()
    reSelectPageGroup
End Sub

Private Sub ListPhotoFills_Click()
    reSelectPhotoFill
End Sub

Private Sub ListRules_Click()
    reSelectRule
End Sub

Private Sub ListSwaps_Click()
    reSelectSwap
End Sub

Private Sub ListLock(ByVal inList As ListBox, ParamArray var() As Variant)
    Dim i As Integer
    For i = LBound(var) To UBound(var)
        LockControl var(i), inList.ListIndex = -1
    Next
End Sub

Private Sub ButtonSave_Click()
    Application.Calculation = xlCalculationManual
    SaveProperties
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub ButtonDupeRule_Click()
    If ListRules.ListIndex <> -1 Then
        Dim newRuleName As String
        newRuleName = Left$(replace(InputBox("Enter New Rule Name:", "Edit Rules"), " ", vbNullString), 100)
        If newRuleName <> vbNullString And collectionGetOrNothing(CurrentProfile.Rules, newRuleName) Is Nothing Then
            Dim ListCases As Collection
            Set ListCases = New Collection
            Dim currentCase As FlnCase
            For Each currentCase In CurrentProfile.Rules(ListRules.ListIndex + 1).Cases
                ListCases.Add caseFromInstantiate(currentCase.Condition, multiRangeFromInstantiate(ThisWorkbook, currentCase.HideRef.ToString), collectionAppendSheets(New Collection, currentCase.Sheets))
            Next
            CurrentProfile.Rules.Add ruleFromInstantiate(newRuleName, ListCases), newRuleName
            DirtyRules
        End If
    End If
End Sub

Private Sub ButtonEditRule_Click()
    If ListRules.ListIndex <> -1 Then
        Dim newRuleName As String
        newRuleName = Left$(replace(InputBox("Enter New Rule Name:", "Edit Rules"), " ", vbNullString), 100)
        If newRuleName <> vbNullString And collectionGetOrNothing(CurrentProfile.Rules, newRuleName) Is Nothing Then
            CurrentProfile.Rules.Add CurrentProfile.Rules(ListRules.ListIndex + 1), newRuleName, After:=ListRules.ListIndex + 1 'Duplicate Reference with new key
            CurrentProfile.Rules.Remove ListRules.ListIndex + 1
            CurrentProfile.Rules(newRuleName).Name = newRuleName
            DirtyRules
        End If
    End If
End Sub

Private Sub ButtonEditCase_Click()
    EditCase
End Sub

Private Sub ButtonEditAssetList_Click()
    If ListAssetLists.ListIndex <> -1 Then
        Dim selectedAssetList As FlnAssetList
        Set selectedAssetList = CurrentProfile.AssetLists(ListAssetLists.ListIndex + 1)
        Me.Hide
        EditAssetList selectedAssetList
        Me.Show
    End If
End Sub

Private Sub ButtonEditCValOutput_Click()
    If ListSwaps.ListIndex <> -1 And ListCVals.ListIndex <> -1 Then
        Dim selectedSwap As FlnSwap
        Dim selectedTable As FlnCValTable
        Set selectedSwap = CurrentProfile.Swaps(ListSwaps.ListIndex + 1)
        Set selectedTable = selectedSwap.CValTables(ListCVals.ListIndex + 1)
        Me.Hide
        EditCValTable selectedSwap, selectedTable
        Me.Show
    End If
End Sub

Private Sub ButtonEditOutput_Click()
    If ListOutputs.ListIndex <> -1 Then
        Dim selectedTable As FlnOutputTable
        Set selectedTable = CurrentProfile.OutputTables(ListOutputs.ListIndex + 1)
        Me.Hide
        EditOutputTable selectedTable
        Me.Show
    End If
End Sub

Private Sub ButtonEditPhotoFill_Click()
    If ListPhotoFills.ListIndex <> -1 Then
        Dim selectedPhotoFill As FlnPhotoFill
        Set selectedPhotoFill = CurrentProfile.photofills(ListPhotoFills.ListIndex + 1)
        Me.Hide
        EditPhotoFill selectedPhotoFill
        Me.Show
    End If
End Sub

Private Sub ButtonEditProfiles_Click()
    EditProfiles
    Me.Show
End Sub

Private Sub ButtonSetAssetCell_Click()
    Me.Hide
    
    Dim CellString As String
    CellString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(CurrentProfile.AssetCell, True, False), "Set Asset Cell", Dimensions:=dimensionsFromRowsCols(1, 1))
    If CellString = vbNullString Then GoTo Err
    
    Set CurrentProfile.AssetCell = rangeOrNothing(CellString, ThisWorkbook)
    DirtyRules
    Me.Show
    Exit Sub
Err:
    MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
    Me.Show
End Sub
Private Sub ButtonClearAssetCell_Click()
    Set CurrentProfile.AssetCell = Nothing
    DirtyRules
End Sub

Private Sub ButtonSetPhotoRename_Click()
    Me.Hide
    
    Dim PhotoRenameString As String
    PhotoRenameString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(CurrentProfile.PhotoRename, True, False), "Set Photo Rename", Dimensions:=dimensionsFromRowsCols(-1, 2))
    If PhotoRenameString = vbNullString Then GoTo Err
    
    Set CurrentProfile.PhotoRename = rangeOrNothing(PhotoRenameString, ThisWorkbook)
    DirtyRules
    Me.Show
    Exit Sub
Err:
    MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
    Me.Show
End Sub
Private Sub ButtonClearPhotoRename_Click()
    Set CurrentProfile.PhotoRename = Nothing
    DirtyRules
End Sub

Private Sub ButtonSetPhotoPath_Click()
    Me.Hide
    
    Dim NewPath As String
    NewPath = pathFromPickerFolder("Pick Photo Folder", "Retrieve photos from here", ThisWorkbook.Path)
    NewPath = replace(NewPath, ThisWorkbook.Path, "%WORKBOOKPATH%")
    If NewPath = vbNullString Then GoTo Err
    CurrentProfile.photoPath = NewPath
    
    DirtyRules
    Me.Show
    Exit Sub
Err:
    MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
    Me.Show
End Sub
Private Sub ButtonClearPhotoPath_Click()
    CurrentProfile.photoPath = vbNullString
    DirtyRules
End Sub

Private Sub ButtonSetFilenameCell_Click()
    Me.Hide
    
    Dim CellString As String
    CellString = stringFromRefEdit(ThisWorkbook, rangeGetAddress(CurrentProfile.FilenameCell, True, False), "Set Filename Cell", Dimensions:=dimensionsFromRowsCols(1, 1))
    If CellString = vbNullString Then GoTo Err
    
    Set CurrentProfile.FilenameCell = rangeOrNothing(CellString, ThisWorkbook)
    DirtyRules
    Me.Show
    Exit Sub
Err:
    MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
    Me.Show
End Sub
Private Sub ButtonClearFilenameCell_Click()
    Set CurrentProfile.FilenameCell = Nothing
    DirtyRules
End Sub

Private Sub ButtonEditSwap_Click()
    If ListSwaps.ListIndex <> -1 Then
        Me.Hide
        EditSwap CurrentProfile.Swaps(ListSwaps.ListIndex + 1)
        Me.Show
    End If
End Sub

Private Sub ButtonEditAutofit_Click()
    If ListAutofits.ListIndex <> -1 Then
        Me.Hide
        EditAutofit CurrentProfile.autofits(ListAutofits.ListIndex + 1)
        Me.Show
    End If
End Sub

Private Sub ButtonEditAutohide_Click()
    If ListAutohides.ListIndex <> -1 Then
        Me.Hide
        EditAutohide CurrentProfile.autohides(ListAutohides.ListIndex + 1)
        Me.Show
    End If
End Sub

Private Sub ButtonEditPageGroup_Click()
    If ListPageGroups.ListIndex <> -1 Then
        Me.Hide
        EditPageGroup CurrentProfile.pageGroups(ListPageGroups.ListIndex + 1)
        Me.Show
    End If
End Sub

Private Sub ButtonNewCase_Click()
    If ListRules.ListIndex <> -1 Then
        Dim selectedRule As FlnRule
        Set selectedRule = CurrentProfile.Rules(ListRules.ListIndex + 1)
        selectedRule.Cases.Add caseFromInstantiate("=FALSE", multiRangeFromInstantiate(ThisWorkbook, inUnionize:=True, inEntireRows:=True), New Collection)
        DirtyRules
    End If
End Sub

Private Sub ButtonNewRule_Click()
    Dim newRuleName As String
    newRuleName = Left$(replace(InputBox("Enter New Rule Name:", "Edit Rules"), " ", vbNullString), 100)
    If newRuleName = vbNullString Or Not collectionGetOrNothing(CurrentProfile.Rules, newRuleName) Is Nothing Then Exit Sub
    
    CurrentProfile.Rules.Add ruleFromInstantiate(newRuleName, New Collection), newRuleName
    DirtyRules
End Sub

Private Sub ButtonNewAssetList_Click()
    Dim newAssetListName As String
    newAssetListName = Left$(replace(InputBox("Enter New Asset List Name:", "Edit Asset Lists"), " ", vbNullString), 100)
    If newAssetListName = vbNullString Or Not collectionGetOrNothing(CurrentProfile.AssetLists, newAssetListName) Is Nothing Then Exit Sub
    
    CurrentProfile.AssetLists.Add assetListFromInstantiate(newAssetListName, ThisWorkbook.Sheets(1).Name & "!" & "A1"), newAssetListName
    DirtyRules
End Sub

Private Sub ButtonNewPhotoFill_Click()
    Dim newPhotoFillName As String
    newPhotoFillName = Left$(replace(InputBox("Enter New Photo Fill Name:", "Edit Photo Fills"), " ", vbNullString), 100)
    If newPhotoFillName = vbNullString Or Not collectionGetOrNothing(CurrentProfile.photofills, newPhotoFillName) Is Nothing Then Exit Sub
    
    Dim newPhotoFill As FlnPhotoFill
    Set newPhotoFill = photoFillFromInstantiate(newPhotoFillName, Nothing, Nothing)
    Me.Hide
    EditPhotoFill newPhotoFill                   ' Will dirty rules on success
    
    If Not newPhotoFill.Dest Is Nothing Then
        CurrentProfile.photofills.Add newPhotoFill
        DirtyRules
    End If
    
    Me.Show
End Sub

Private Sub ButtonNewCVal_Click()
    If ListSwaps.ListIndex <> -1 Then
        Dim selectedSwap As FlnSwap
        Set selectedSwap = CurrentProfile.Swaps(ListSwaps.ListIndex + 1)
        Dim newTableName As String
        newTableName = Left$(replace(InputBox("Enter New CVal Table Name:", "Edit CVal Tables"), " ", vbNullString), 100)
        If newTableName = vbNullString Or Not collectionGetOrNothing(selectedSwap.CValTables, newTableName) Is Nothing Then Exit Sub
        Dim newTable As FlnCValTable
        Set newTable = cValTableFromInstantiate(newTableName, Nothing, Nothing)
        Me.Hide
        EditCValTable selectedSwap, newTable
        If Not newTable.Dest Is Nothing Then
            selectedSwap.CValTables.Add newTable, newTable.Name
            DirtyRules
        End If
        Me.Show
    End If
End Sub

Private Sub ButtonNewOutput_Click()
    Dim newTableName As String
    newTableName = Left$(replace(InputBox("Enter New Output Table Name:", "Edit Output Tables"), " ", vbNullString), 100)
    If newTableName = vbNullString Or Not collectionGetOrNothing(CurrentProfile.OutputTables, newTableName) Is Nothing Then Exit Sub
    Dim newTable As FlnOutputTable
    Set newTable = outputTableFromInstantiate(newTableName, Nothing, Nothing, Nothing)
    Me.Hide
    EditOutputTable newTable
    If Not newTable.Dest Is Nothing Then
        CurrentProfile.OutputTables.Add newTable, newTable.Name
        DirtyRules
    End If
    Me.Show
End Sub

Private Sub ButtonNewSwap_Click()
    Dim newSwapName As String
    newSwapName = Left$(replace(InputBox("Enter New Swap Name:", "Edit Swaps"), " ", vbNullString), 100)
    If newSwapName = vbNullString Or Not collectionGetOrNothing(CurrentProfile.Swaps, newSwapName) Is Nothing Then Exit Sub
    
    Dim newSwap As FlnSwap
    Set newSwap = swapFromInstantiate(newSwapName, False, Nothing, Nothing, vbNullString, Nothing, New Collection, New Collection, New Collection)
    Me.Hide
    EditSwap newSwap
    If Not newSwap.SwapCell Is Nothing Then
        CurrentProfile.Swaps.Add newSwap, newSwap.Name
        DirtyRules
    End If
    Me.Show
End Sub

Private Sub ButtonNewDupe_Click()
    If ListSwaps.ListIndex <> -1 Then
        Dim selectedSwap As FlnSwap
        Set selectedSwap = CurrentProfile.Swaps(ListSwaps.ListIndex + 1)
        
        With New SvnPicker
            .Initialize "Add Dupe Sheet", sheetsToNames(collectionRemoveSheets(workbookGetSheets(ThisWorkbook), selectedSwap.DupeSheets))
            .Show
            If .Success Then
                selectedSwap.DupeSheets.Add ThisWorkbook.Sheets(.Text), .Text
                selectedSwap.DoDupeSplit.Add False, .Text
                DirtyRules
            End If
        End With
    End If
End Sub

Private Sub ButtonNewAutofit_Click()
    Dim newAutofitName As String
    newAutofitName = Left$(replace(InputBox("Enter New Autofit Name:", "Edit Autofits"), " ", vbNullString), 100)
    If newAutofitName = vbNullString Or Not collectionGetOrNothing(CurrentProfile.autofits, newAutofitName) Is Nothing Then Exit Sub
    
    Dim newAutofit As FlnAutoFit
    Set newAutofit = autoFitFromInstantiate(newAutofitName, vbNullString)
    Me.Hide
    EditAutofit newAutofit
    
    If Not newAutofit.MultiRange.ToString = vbNullString Then
        CurrentProfile.autofits.Add newAutofit
        DirtyRules
    End If
    
    Me.Show
End Sub

Private Sub ButtonNewAutohide_Click()
    Dim newAutohideName As String
    newAutohideName = Left$(replace(InputBox("Enter New Autohide Name:", "Edit Autohides"), " ", vbNullString), 100)
    If newAutohideName = vbNullString Or Not collectionGetOrNothing(CurrentProfile.autohides, newAutohideName) Is Nothing Then Exit Sub
    
    Dim newAutohide As FlnAutoHide
    Set newAutohide = autoHideFromInstantiate(newAutohideName, vbNullString)
    Me.Hide
    EditAutohide newAutohide
    
    If Not newAutohide.MultiRange.ToString = vbNullString Then
        CurrentProfile.autohides.Add newAutohide
        DirtyRules
    End If
    
    Me.Show
End Sub

Private Sub ButtonNewPageGroup_Click()
    Dim newPageGroupName As String
    newPageGroupName = Left$(replace(InputBox("Enter New PageGroup Name:", "Edit PageGroups"), " ", vbNullString), 100)
    If newPageGroupName = vbNullString Or Not collectionGetOrNothing(CurrentProfile.pageGroups, newPageGroupName) Is Nothing Then Exit Sub
    
    Dim newPageGroup As FlnPageGroup
    Set newPageGroup = pageGroupFromInstantiate(newPageGroupName, vbNullString)
    Me.Hide
    EditPageGroup newPageGroup
    
    If Not newPageGroup.MultiRange.ToString = vbNullString Then
        CurrentProfile.pageGroups.Add newPageGroup
        DirtyRules
    End If
    Me.Show
End Sub

Private Sub ButtonDeleteAutofit_Click()
    If ListAutofits.ListIndex <> -1 Then
        CurrentProfile.autofits.Remove ListAutofits.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonDeleteAutohide_Click()
    If ListAutohides.ListIndex <> -1 Then
        CurrentProfile.autohides.Remove ListAutohides.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonRemovePageGroup_Click()
    If ListPageGroups.ListIndex <> -1 Then
        CurrentProfile.pageGroups.Remove ListPageGroups.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonRemoveRule_Click()
    If ListRules.ListIndex <> -1 And MsgBox("This will erase all cases! Are you sure?", vbYesNo, "Delete Rule") = vbYes Then
        CurrentProfile.Rules.Remove ListRules.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonRemoveCase_Click()
    If ListRules.ListIndex <> -1 And ListCases.ListIndex <> -1 Then
        Dim selectedRule As FlnRule
        Set selectedRule = CurrentProfile.Rules(ListRules.ListIndex + 1)
        selectedRule.Cases.Remove ListCases.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonRemoveAssetList_Click()
    If ListAssetLists.ListIndex <> -1 Then
        CurrentProfile.AssetLists.Remove ListAssetLists.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonRemoveCVal_Click()
    If ListSwaps.ListIndex <> -1 And ListCVals.ListIndex <> -1 Then
        CurrentProfile.Swaps(ListSwaps.ListIndex + 1).CValTables.Remove ListCVals.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonRemoveOutput_Click()
    If ListOutputs.ListIndex <> -1 Then
        CurrentProfile.OutputTables.Remove ListOutputs.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonRemoveDupe_Click()
    If ListDupes.ListIndex <> -1 And ListSwaps.ListIndex <> -1 Then
        CurrentProfile.Swaps(ListSwaps.ListIndex + 1).DupeSheets.Remove ListDupes.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonRemoveSwap_Click()
    If ListSwaps.ListIndex <> -1 Then
        CurrentProfile.Swaps.Remove ListSwaps.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonRemovePhotoFill_Click()
    If ListPhotoFills.ListIndex <> -1 Then
        CurrentProfile.photofills.Remove ListPhotoFills.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ToggleDupeSplit_Change()
    Dim selectedSwap As FlnSwap
    If ListSwaps.ListIndex <> -1 And ListDupes.ListIndex <> -1 And ToggleDupeSplit.Tag <> "Locked" Then
        Set selectedSwap = CurrentProfile.Swaps(ListSwaps.ListIndex + 1)
        If ListDupes.ListIndex + 1 = selectedSwap.DupeSheets.count Then
            collectionSetByKey selectedSwap.DoDupeSplit, Not selectedSwap.DoDupeSplit(ListDupes.ListIndex + 1), selectedSwap.DupeSheets(ListDupes.ListIndex + 1).Name
        Else
            collectionSetByKey selectedSwap.DoDupeSplit, Not selectedSwap.DoDupeSplit(ListDupes.ListIndex + 1), selectedSwap.DupeSheets(ListDupes.ListIndex + 1).Name, selectedSwap.DupeSheets(ListDupes.ListIndex + 1).Name
        End If
        ToggleSwapEnd.Tag = "Locked"
        DirtyRules
        ToggleSwapEnd.Tag = vbNullString
    End If
End Sub

Private Sub ToggleSwapEnd_Change()
    Dim selectedSwap As FlnSwap
    If ListSwaps.ListIndex <> -1 And ToggleSwapEnd.Tag <> "Locked" Then
        Set selectedSwap = CurrentProfile.Swaps(ListSwaps.ListIndex + 1)
        selectedSwap.doEndOrder = Not selectedSwap.doEndOrder
        ToggleSwapEnd.Tag = "Locked"
        DirtyRules
        ToggleSwapEnd.Tag = vbNullString
    End If
End Sub

Private Sub ButtonCaseDown_Click()
    If ListRules.ListIndex <> -1 And ListCases.ListIndex <> -1 Then reOrder ListCases, CurrentProfile.Rules(ListRules.ListIndex + 1).Cases, ListCases.ListIndex + 1, False
End Sub

Private Sub ButtonCaseUp_Click()
    If ListRules.ListIndex <> -1 And ListCases.ListIndex <> -1 Then reOrder ListCases, CurrentProfile.Rules(ListRules.ListIndex + 1).Cases, ListCases.ListIndex + 1, True
End Sub

Private Sub ButtonRuleDown_Click()
    If ListRules.ListIndex <> -1 Then reOrder ListRules, CurrentProfile.Rules, ListRules.ListIndex + 1, False
End Sub

Private Sub ButtonRuleUp_Click()
    If ListRules.ListIndex <> -1 Then reOrder ListRules, CurrentProfile.Rules, ListRules.ListIndex + 1, True
End Sub

Private Sub PageSwitcher_Change()
    If PageSwitcher.value = 0 Then
        Me.Height = 36 + FrameGeneratorTools.Top + FrameGeneratorTools.Height + 24
        Me.Width = FrameGeneratorTools.Left + FrameGeneratorTools.Width + 24
    Else
        Me.Height = PageSwitcher.Height + 66
        Me.Width = PageSwitcher.Width
    End If
End Sub

''' General Events '''

Private Sub UserForm_Initialize()                ' On form load
    ResetForm                                    ' Reset form to default state
    ' Load Properties
    LoadProperties
    reLoadProfiles
    PageSwitcher_Change
    Exit Sub
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) ' On form attempted Close
    If LabelSaveWarning.Visible Then
        Dim response As Long
        response = MsgBox("You've made unsaved changes to one or more profiles." + vbNewLine + "Save these changes before exiting?", vbYesNoCancel + vbQuestion, "Warning: Potential data loss")
        If response = vbYes Then
            SaveProperties
        ElseIf response <> vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Application.EnableCancelKey = xlInterrupt    ' Reset Break Key to default
End Sub


