VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlnGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

' Configuration Constants
Const PHOTO_TITLE As String = "FLN_PHOTOFILL"
Const DUPE_PROPERTY As String = "FLN_DUPLICATE"
Private RuleProfile As FlnProfile
Private SetAssets As FlnAssetList
Private OutputTables As Collection
Private AssetCell As Range
Private PhotoRename As Range
Private FilenameCell As Range
Private Swaps As Collection
Private photofills As Collection
Private autofits As Collection
Private autohides As Collection
Private pageGroups As Collection
Private photoPath As String
Private fso As FileSystemObject
Private logOutputPath As String
Private logOutput As TextStream

Private pDebug As Boolean
    
' Working Variables
Private returnSheet As Worksheet
Private copiedSheets As Collection
Private returnEID As Variant
Private exportMode As Boolean
Private FolderPath As String
Private AssetSheets As Collection
Private SplitDupeSheets As Collection
Private SwapCombos As Collection
Private DupeIndexes As Collection
Private ETA As String
Private addedShapes As Collection
Private addedBreaks As Collection
Private LastSvnResult As SvnResult
Private assetsCompleted As Long
Private currentComboIndex As Long

' Looping Variables
Private currentCombo As Collection
Private currentID As Variant
Private currentRow As Long
Private comboIndex As Long
Private currentSwap As FlnSwap
Private currentAutohide As FlnAutoHide
Private currentCVal As FlnCValTable
Private currentOutTable As FlnOutputTable
Private currentPhotoFill As FlnPhotoFill
Private currentPhotoShape As Shape
Private currentBreakRange As Range
Private currentPhotoArea As Range

Public Sub Run()

    ' Set up output log
    Dim logFolderPath As String
    Set fso = CreateObject("scripting.filesystemobject")
    logFolderPath = CreateObject("Wscript.Shell").SpecialFolders("MyDocuments") & "\FlorinLogs"
    If Not fso.FolderExists(logFolderPath) Then fso.CreateFolder logFolderPath
    logOutputPath = logFolderPath & "\" & RuleProfile.Name & " (" & Format$(Now(), "yyyy-mm-dd-hh-mm-ss") & ")" & ".txt"
    Set logOutput = fso.CreateTextFile(logOutputPath)
    
    ' Speedups and cancel-handling
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set copiedSheets = New Collection
    On Error GoTo Handler
    If Not pDebug Then Application.EnableCancelKey = xlErrorHandler Else Application.EnableCancelKey = xlInterrupt

    FlnUI.SetStatus logOutput, inLine1:="Generating '" & SetAssets.Name & "' using profile " & RuleProfile.Name & ". Cancel using CTRL+Break.", inLine2:="Performing initial setup..."
    
    ' Save initial EID and Active Sheet to restore later
    returnEID = AssetCell.Formula
    Set returnSheet = ActiveSheet
    
    FlnUI.SetStatus logOutput, inLine3:="Checking photo directory..."
    
    ' Check if photofill directories are valid
    If photoPath <> vbNullString And SafeDir(replace(photoPath, "%WORKBOOKPATH%", ThisWorkbook.Path), vbDirectory) = vbNullString Then
        If MsgBox("Expected photofill path not found (or is online): " & replace(photoPath, "%WORKBOOKPATH%", ThisWorkbook.Path) & vbCrLf & "Disable photos and generate reports anyway?", vbOKCancel) = vbOK Then
            Set photofills = New Collection
            photoPath = vbNullString
        Else
            GoTo CancelCleanup
        End If
    End If
        
    ' Check Export Mode
    exportMode = RuleProfile.GetSheets.count > 0
    
    FlnUI.SetStatus logOutput, inLine3:="Picking export folder..."
    
    ' Pick export folder
    If (exportMode Or Not PhotoRename Is Nothing) Then
        FolderPath = pickReportFolder()
        If (FolderPath = vbNullString) Then GoTo CancelCleanup
    End If
    
    FlnUI.SetStatus logOutput, inLine3:="Setting up OutputTable algorithm..."
    
    ' Set up Output Algorithm
    For Each currentOutTable In OutputTables
        currentOutTable.SetupAlgo
    Next
    If (exportMode) Then GetDupeIndexes
    
    Dim startTime As Double
    startTime = Timer                            ' ETA start time
    
    ''' Calculation and Generation Cycle '''
            
    ' For Each Asset
    For Each currentID In SetAssets.MultiRange.Areas(1).Rows
        FlnUI.SetStatus logOutput, inLine2:=getETA("Generating Report", startTime, (currentID.row - SetAssets.MultiRange.Areas(1).row) + 1, SetAssets.MultiRange.Areas(1).Rows.count), inLine3:=vbNullString
        
        DoEvents
        
        Set AssetSheets = New Collection         ' Reset asset sheets
        Set SplitDupeSheets = New Collection     ' Individual PDF prints
        Set addedShapes = New Collection         ' Reset (now deleted) photos
        Set addedBreaks = New Collection
        
        ' Set new Asset ID
        AssetCell.Value2 = currentID.Value2
        
        FlnUI.SetStatus logOutput, inLine3:="Recalculating formulae..."
        
        Application.EnableCancelKey = xlDisabled
        On Error GoTo 0                          ' Don't Catch UDF Cell Errors
        Application.Calculate
        On Error GoTo Handler
        If Not pDebug Then Application.EnableCancelKey = xlErrorHandler Else Application.EnableCancelKey = xlInterrupt
        
        If Swaps.count > 0 Then FlnUI.SetStatus logOutput, inLine3:="Validating and clearing Swaps..."
        
        ' Validate and Clear Swaps
        For Each currentSwap In Swaps
            For Each currentCVal In currentSwap.CValTables
                If (Not currentSwap.swapRange Is Nothing) Then ' Ensure Range Exists
                    If (currentSwap.swapRange.count > currentCVal.Dest.Areas.count + 1) Then
                        If errCancel(IIf(IsEmpty(currentID), "Invalid (Empty) EID", "Invalid Swapset Size"), CStr(currentSwap.swapRange.count) & " Swaps(s)") Then GoTo Handler
                        GoTo NextAsset
                    End If
                End If
                currentCVal.Dest.ClearContents
            Next
        Next
        
        
        FlnUI.SetStatus logOutput, inLine3:="Calculating Swap combinations..."
        ' Calculate Swap Combos
        On Error GoTo 0                          ' Defer to sub error handling
        Set LastSvnResult = GetSwapCombos
        On Error GoTo Handler
        If Not pDebug Then Application.EnableCancelKey = xlErrorHandler Else Application.EnableCancelKey = xlInterrupt
        If LastSvnResult.ResultType = svnSkip Then GoTo NextAsset
        If LastSvnResult.ResultType = svnCancel Then GoTo Handler
        If LastSvnResult.ResultType = SvnSubErr Then Err.Raise LastSvnResult.ErrorNum, Description:=LastSvnResult.ErrorDesc
        
        
        FlnUI.SetStatus logOutput, inLine3:="Hiding PhotoFill destinations..."
        For Each currentPhotoFill In photofills
            For Each currentPhotoArea In currentPhotoFill.Dest.Areas
                If currentPhotoArea.MergeCells Then currentPhotoArea.MergeArea.EntireRow.Hidden = True
                If Not currentPhotoArea.MergeCells Then currentPhotoArea.EntireRow.Hidden = True
            Next
        Next
        
        For currentComboIndex = 1 To SwapCombos.count
            Set currentCombo = SwapCombos(currentComboIndex)
            FlnUI.SetStatus logOutput, inLine3:="Performing swapwise generation (" & currentComboIndex & "/" & SwapCombos.count & ")"
            DoEvents
        
        
            FlnUI.SetStatus logOutput, inLine4:="Performing recursive formulae calculation..."
            ' Upate Swap and Validate swap Values
            For comboIndex = 1 To currentCombo.count
                GetRelevantSwap(comboIndex).SwapCell.Value2 = currentCombo(comboIndex)
                If IsEmpty(GetSwapIndex(comboIndex)) Then
                    If errCancel("Invalid Swapset Value", "Value " & currentCombo(comboIndex)) Then GoTo Handler
                    GoTo NextAsset
                End If
                
                ' Force recalculate
                Application.EnableCancelKey = xlDisabled
                On Error GoTo 0                  ' Don't Catch UDF Cell Errors
                Application.Calculate
                On Error GoTo Handler
                If Not pDebug Then Application.EnableCancelKey = xlErrorHandler Else Application.EnableCancelKey = xlInterrupt
                
                DoEvents
            Next
            
            If (exportMode) Then
                FlnUI.SetStatus logOutput, inLine4:="Adding sheets to export..."
                ''' Add Sheets For Printing '''
                collectionAppendSheets AssetSheets, getSwapSheets(RuleProfile.getActiveSheets)
                
                If photofills.count > 0 Then FlnUI.SetStatus logOutput, inLine4:="Inserting photos..."
                ' Add Photos
                Set LastSvnResult = AddPhotos
                If LastSvnResult.ResultType = svnSkip Then GoTo NextAsset
                If LastSvnResult.ResultType = svnCancel Then GoTo Handler
                
                FlnUI.SetStatus logOutput, inLine4:="Duplicating sheets..."
                ' Split duplicates
                For Each currentSwap In Swaps
                    Dim currentDupe As Worksheet
                    For Each currentDupe In currentSwap.DupeSheets
                        If currentSwap.DoDupeSplit(currentDupe.Name) Then
                            Dim copyFriendlySheet As Worksheet
                            Set copyFriendlySheet = CopyFriendlyGet(currentDupe.Name, False)
                            If collectionGetOrNothing(SplitDupeSheets, copyFriendlySheet.Name) Is Nothing Then SplitDupeSheets.Add copyFriendlySheet, copyFriendlySheet.Name
                            If Not collectionGetOrNothing(AssetSheets, copyFriendlySheet.Name) Is Nothing Then AssetSheets.Remove copyFriendlySheet.Name
                        End If
                    Next
                Next
                
                
                FlnUI.SetStatus logOutput, inLine4:="Performing prettify..."
                ''' Autofit Ranges '''
                Set LastSvnResult = applyAutofit
                If LastSvnResult.ResultType = svnSkip Then GoTo NextAsset
                If LastSvnResult.ResultType = svnCancel Then GoTo Handler
                
                ''' Autohide Ranges '''
                Set LastSvnResult = applyAutohide
                If LastSvnResult.ResultType = svnSkip Then GoTo NextAsset
                If LastSvnResult.ResultType = svnCancel Then GoTo Handler
                
                ''' Show/Hide Ranges For Printing '''
                toggleAreaVisibility RuleProfile.getHideRange, False
                toggleAreaVisibility RuleProfile.getActiveHide, True
                
                ''' Auto Page Break '''
                Set LastSvnResult = applyPageGroups
                If LastSvnResult.ResultType = svnSkip Then GoTo NextAsset
                If LastSvnResult.ResultType = svnCancel Then GoTo Handler
                
            End If
            
            FlnUI.SetStatus logOutput, inLine4:="Freezing CValTables..."
            ' Freeze CVals
            For comboIndex = currentCombo.count To 1 Step -1 ' In reverse so nesting behaviour works as expected
                If GetSwapIndex(comboIndex) > 1 Then
                    For Each currentCVal In GetRelevantSwap(comboIndex).CValTables
                        ' Freeze to current values
                        currentCVal.Dest.Areas(GetSwapIndex(comboIndex) - 1).Value2 = currentCVal.Source.Value2
                    Next
                End If
            Next
        Next
        
        FlnUI.SetStatus logOutput, inLine3:="Performing non-swap generation...", inLine4:="Resetting swaps"
        
        ' Reset Swaps to Base to ensure valid active cvals
        For Each currentSwap In Swaps
            currentSwap.SwapCell.Value2 = currentSwap.MaxSwapSet.Rows(1).Value2
        Next
        
        ' Double check sheets and ranges in case no swaps
        If (SwapCombos.count = 0 And exportMode) Then
            FlnUI.SetStatus logOutput, inLine4:="Adding sheets to export..."
            ''' Add Sheets For Printing '''
            collectionAppendSheets AssetSheets, RuleProfile.getActiveSheets
            
            If photofills.count > 0 Then FlnUI.SetStatus logOutput, inLine4:="Inserting photos..."
            ' Add Photos
            Set LastSvnResult = AddPhotos(False)
            If LastSvnResult.ResultType = svnSkip Then GoTo NextAsset
            If LastSvnResult.ResultType = svnCancel Then GoTo Handler
            
            FlnUI.SetStatus logOutput, inLine4:="Performing prettify..."
            ''' Autofit Ranges '''
            Set LastSvnResult = applyAutofit(False)
            If LastSvnResult.ResultType = svnSkip Then GoTo NextAsset
            If LastSvnResult.ResultType = svnCancel Then GoTo Handler
            
            ''' Autohide Ranges '''
            Set LastSvnResult = applyAutohide(False)
            If LastSvnResult.ResultType = svnSkip Then GoTo NextAsset
            If LastSvnResult.ResultType = svnCancel Then GoTo Handler
            
            ''' Show/Hide Ranges For Printing '''
            toggleAreaVisibility RuleProfile.getHideRange, False, False
            toggleAreaVisibility RuleProfile.getActiveHide, True, False
                
            ''' Auto Page Break '''
            Set LastSvnResult = applyPageGroups(False)
            If LastSvnResult.ResultType = svnSkip Then GoTo NextAsset
            If LastSvnResult.ResultType = svnCancel Then GoTo Handler
        End If
        
        FlnUI.SetStatus logOutput, inLine4:="Recalculating formulae..."
        
        ' Recalculate anything that depends on frozen cvals
        Application.EnableCancelKey = xlDisabled
        On Error GoTo 0                          ' Don't Catch UDF Cell Errors
        Application.Calculate
        On Error GoTo Handler
        If Not pDebug Then Application.EnableCancelKey = xlErrorHandler Else Application.EnableCancelKey = xlInterrupt
        
        
        DoEvents
        
        If OutputTables.count > 0 Then FlnUI.SetStatus logOutput, inLine3:="Populating output tables...", inLine4:=vbNullString
        ' Freeze asset output
        Dim RangeToDelete As Range
        For Each currentOutTable In OutputTables
            For currentRow = 1 To currentOutTable.Source.Rows.count
                If IsError(currentOutTable.Source.Areas(3).Rows(currentRow).Value2) Then ' Output Garbage Rows for clarity
                    AlgoGetRow(currentID.Value2, currentOutTable.Source.Areas(1).Rows(currentRow), currentOutTable).Value2 = currentOutTable.Source.Areas(2).Rows(currentRow).Value2
                ElseIf currentOutTable.Source.Areas(3).Rows(currentRow).Value2 <> False Then ' Inserts if doesn't exist
                    AlgoGetRow(currentID.Value2, currentOutTable.Source.Areas(1).Rows(currentRow), currentOutTable).Value2 = currentOutTable.Source.Areas(2).Rows(currentRow).Value2
                Else                             ' Delete Existing False Rows
                    Set RangeToDelete = GetRow(currentID.Value2, currentOutTable.Source.Areas(1).Rows(currentRow), currentOutTable)
                    If Not RangeToDelete Is Nothing Then RangeToDelete.EntireRow.Delete
                End If
            Next
        Next
            
        If (exportMode) Then
            FlnUI.SetStatus logOutput, inLine3:="Exporting...", inLine4:="Exporting sheets to PDF..."
            ' Export
            Dim FilePath As String
            FilePath = FolderPath & "\" & FilenameCell.Value2
            DoEvents
            exportSheetsToPath FilePath, AssetSheets, SplitDupeSheets
            DoEvents
            
            
            ' XXX - Copy Photos
            If Not RuleProfile.PhotoRename Is Nothing And photoPath <> vbNullString Then
                FlnUI.SetStatus logOutput, inLine4:="Copying and renaming photos..."
                MkDir FolderPath & "\" & FilenameCell.Value2
                Dim PhotoRenameArray As Variant
                PhotoRenameArray = RuleProfile.PhotoRename.Value2
                For currentRow = LBound(PhotoRenameArray, 1) To UBound(PhotoRenameArray, 1)
                    If PhotoRenameArray(currentRow, 1) <> vbNullString Then
                        Dim OldFileName As String
                        Dim ext As String
                        OldFileName = replace(photoPath, "%WORKBOOKPATH%", ThisWorkbook.Path) & "\" & PhotoRenameArray(currentRow, 1)
                        If Dir(OldFileName & ".jpg") <> vbNullString Or Dir(OldFileName & ".png") <> vbNullString Then
                            ext = IIf(Dir(OldFileName & ".jpg") <> vbNullString, ".jpg", ".png")
                            fso.CopyFile OldFileName & ext, FolderPath & "\" & FilenameCell.Value2 & "\" & PhotoRenameArray(currentRow, 2) & ext
                        End If
                    End If
                Next
            End If
        End If
        
        FlnUI.SetStatus logOutput, inLine3:="Preparing for next asset...", inLine4:=vbNullString
        ' Delete Photos
        For Each currentPhotoShape In addedShapes
            currentPhotoShape.Delete
        Next
        Set addedShapes = New Collection         ' In case of early cleanup
        
        assetsCompleted = assetsCompleted + 1
NextAsset:
    Next

CleanUp:
    ''' Cleanup '''
    Application.EnableCancelKey = xlDisabled     'Prevent Double-Error
    On Error GoTo 0                              ' Reset Error Handling
    FlnUI.SetStatus logOutput, inLine1:="Finished Generating " & assetsCompleted & "/" & SetAssets.MultiRange.Areas(1).Rows.count & " from '" & SetAssets.Name & "' using profile " & RuleProfile.Name, inLine2:=vbNullString, inLine3:=vbNullString, inLine4:=vbNullString
    CleanUp
    
    ' Whether to open explorer after successful export
    If FolderPath <> vbNullString Then
        If MsgBox("Open the generated files in Explorer?", vbYesNo, "Export Successful") = vbYes Then SvnOffice.openInExplorer FolderPath
    End If
    
    Exit Sub
    
CancelCleanup:
    Application.EnableCancelKey = xlDisabled     'Prevent Double-Error
    On Error GoTo 0                              ' Reset Error Handling
    FlnUI.SetStatus logOutput, inLine1:="Stopped after generating " & assetsCompleted & "/" & SetAssets.MultiRange.Areas(1).Rows.count & " from '" & SetAssets.Name & "' using profile " & RuleProfile.Name & ":"
    CleanUp
    Exit Sub
    
Handler:
    Application.EnableCancelKey = xlDisabled     'Prevent Double-Error
    
    If (pDebug) Then
        If (HandleError(Err)) = vbYes Then
            Stop
            Resume
        End If
    ElseIf (Err.Number <> 0 And Err.Number <> 18) Then
        With New SvnMsgBox
            .Initialize "Error " & CStr(Err.Number) & " occurred during generation", "Message: " & Err.Description & vbNewLine & "Any files created before the error are still saved on your PC.", _
                IIf(FolderPath <> vbNullString, "Delete Files", vbNullString), _
                IIf(FolderPath <> vbNullString, "Open Folder", vbNullString), "Continue"
            .Show vbModal
            ' Handle which button was pressed
            If (.ReturnType = svnLeft) Then fso.DeleteFolder FolderPath
            If (.ReturnType = svnCentre) Then SvnOffice.openInExplorer FolderPath
        End With
        If MsgBox("View the log file?", vbYesNo, "Error Occurred") = vbYes Then SvnOffice.openInExplorer logOutputPath
    Else
        With New SvnMsgBox
            .Initialize "Export Cancelled", "Any files created before the error are still saved on your PC.", _
                IIf(FolderPath <> vbNullString, "Delete Files", vbNullString), _
                IIf(FolderPath <> vbNullString, "Open Folder", vbNullString), "Continue"
            .Show vbModal
            ' Handle which button was pressed
            If (.ReturnType = svnLeft) Then fso.DeleteFolder FolderPath
            If (.ReturnType = svnCentre) Then SvnOffice.openInExplorer FolderPath
        End With
    End If
    Resume CancelCleanup
End Sub

Public Sub CleanUp()
    FlnUI.SetStatus logOutput, inLine2:="Cleaning up..."


    FlnUI.SetStatus logOutput, inLine3:="Removing duplicate sheets..."
        ''' Remove Duplicate Sheets '''
    Application.DisplayAlerts = False        ' Disable "Are you sure?" messages
    Dim sheetToDelete As Variant
    For Each sheetToDelete In workbookGetSheets(ThisWorkbook)
        If Not IsEmpty(collectionGetPropertyValueOrEmpty(worksheetGetProperties(sheetToDelete), DUPE_PROPERTY)) Then sheetToDelete.Delete
    Next
    Application.DisplayAlerts = True
    
    FlnUI.SetStatus logOutput, inLine3:="Unprettifying..."
    ''' Re-Show all hidden ranges '''
    If Not RuleProfile Is Nothing Then toggleAreaVisibility RuleProfile.getHideRange, False, False
    
    For Each currentAutohide In autohides ' Show autohide ranges
        toggleAreaVisibility currentAutohide.MultiRange, False, False
    Next
    
    FlnUI.SetStatus logOutput, inLine3:="Restoring Swaps and wiping CValTables..."
    ' Restore Swaps
    For Each currentSwap In Swaps
        currentSwap.SwapCell.Value2 = currentSwap.MaxSwapSet.Rows(1).Value2
        For Each currentCVal In currentSwap.CValTables ' Clear CVals
            currentCVal.Dest.ClearContents
        Next
    Next
    
    If Not AssetCell Is Nothing And Not IsEmpty(returnEID) Then AssetCell.Formula = returnEID ' Restore Asset
    
    FlnUI.SetStatus logOutput, inLine3:="Unprettifying PhotoFills..."
    ' Cleanup Photo Area
    For Each currentPhotoFill In photofills
        For Each currentPhotoArea In currentPhotoFill.Dest.Areas
            currentPhotoArea.EntireRow.Hidden = False
        Next
    Next
    
    FlnUI.SetStatus logOutput, inLine3:="Removing photos..."
    ' Delete ALL Photos
    Dim curShape As Shape
    Dim currentSheet As Worksheet
    For Each currentSheet In ThisWorkbook.Sheets
        For Each curShape In currentSheet.Shapes
            If curShape.Type = msoPicture Then
                If curShape.title = PHOTO_TITLE Then curShape.Delete
            End If
        Next
    Next
    
    FlnUI.SetStatus logOutput, inLine3:="Unprettifying page breaks..."
    ' Delete Page breaks
    If Not addedBreaks Is Nothing Then
        For Each currentBreakRange In addedBreaks
            currentBreakRange.PageBreak = xlPageBreakNone
        Next
    End If
    
    FlnUI.SetStatus logOutput, inLine3:="Changing settings back to default..."
    ' Undo speedups
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    
    FlnUI.SetStatus logOutput, inLine3:="Closing log file..."
    If Not logOutput Is Nothing Then logOutput.Close
    
    ' Reset form
    FlnUI.SetStatus Nothing, inLine2:="Ready", inLine3:=vbNullString, inLine4:=vbNullString
    FlnUI.ResetForm
End Sub

Private Function getETA(ByVal Action As String, ByVal startTime As Long, ByVal currentJob As Long, ByVal totalJobs As Long) As String
    Dim minutesElapsed As Double
    Dim ETA As Double
    
    minutesElapsed = (Timer - startTime) / 60
    ' AVG Time taken per report * Remaining Reports, rounded to the nearest 1/2 minute
    ETA = Round((minutesElapsed / currentJob) * (totalJobs - currentJob) * 2, 0) / 2
    getETA = Action & " " _
           & currentJob & " of " _
           & totalJobs _
           & " (ETA: " _
           & IIf(currentJob > totalJobs / 17, ETA & " Minutes Remaining)", "Calculating...)")
    ' Displays after 1/17 (6%) complete.
End Function

Private Function AlgoGetRow(ByVal ID As Variant, ByVal Sigs As Range, ByVal OutTable As FlnOutputTable) As Range
    Set AlgoGetRow = GetRow(ID, Sigs, OutTable)
    If AlgoGetRow Is Nothing Then
        Dim foundRow As Boolean
        Dim currentRow As Long
        For currentRow = 1 To OutTable.Dest.Rows.count
            foundRow = False
            If ID = OutTable.Dest.Areas(1).Rows(currentRow).Value2 Then ' Nesting Required for invalid "END" sig
                If OutTable.GetSigRow(OutTable.Dest.Areas(2).Rows(currentRow)) > OutTable.GetSigRow(Sigs) Then foundRow = True
            End If
            If OutTable.GetIDRow(OutTable.Dest.Areas(1).Rows(currentRow).Value2) > OutTable.GetIDRow(ID) Or currentRow = OutTable.Dest.Rows.count Then foundRow = True ' Insert Above and Return
            If foundRow Then
                OutTable.Dest.EntireRow(currentRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                OutTable.Dest.Areas(1).Rows(currentRow).Offset(0).Value2 = ID
                OutTable.Dest.Areas(2).Rows(currentRow).Offset(0).Value2 = Sigs.Value2
                Set AlgoGetRow = OutTable.Dest.Areas(3).Rows(currentRow).Offset(0)
                Exit Function
            End If
            
        Next
    End If
End Function

Private Function GetRow(ByVal ID As Variant, ByVal Sigs As Range, ByVal OutTable As FlnOutputTable) As Range
    Dim currentRow As Long
    For currentRow = 1 To OutTable.Dest.Rows.count
        If ID & rangeGetText(Sigs) = OutTable.Dest.Areas(1).Rows(currentRow).Value2 & rangeGetText(OutTable.Dest.Areas(2).Rows(currentRow)) Then
            Set GetRow = OutTable.Dest.Areas(3).Rows(currentRow)
            Exit For
        End If
    Next
End Function

Private Sub exportSheetsToPath(ByRef FilePath As String, ByVal inSheets As Collection, ByVal individualSheets As Collection)
    ''' Prints a given set of sheet names from the current workbook. ''
    ' Outputs to filePath
    Dim sheetNames As Variant
    If inSheets.count > 0 Then
        sheetNames = collectionGetSheetVariant(inSheets)
        ThisWorkbook.Sheets(sheetNames).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                        FileName:=FilePath & ".pdf", _
                                        Quality:=xlQualityStandard, _
                                        IncludeDocProperties:=True, _
                                        IgnorePrintAreas:=False, _
                                        OpenAfterPublish:=False
        returnSheet.Select
    End If
    Dim currentSheet As Worksheet
    Dim count As Long
    For Each currentSheet In individualSheets
        count = count + 1
        currentSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                         FileName:=FilePath & "_" & CStr(count) & ".pdf", _
                                         Quality:=xlQualityStandard, _
                                         IncludeDocProperties:=True, _
                                         IgnorePrintAreas:=False, _
                                         OpenAfterPublish:=False
    Next
    returnSheet.Select
End Sub

Private Function pickReportFolder() As String
    ''' Get Directory path using OS directory picker '''
    Dim pickedFolder As String
    pickedFolder = pathFromPickerFolder("Save Reports to Where?", "Generate Report Folder Here", ThisWorkbook.Path)
    If pickedFolder = vbNullString Or Not fso.FolderExists(pickedFolder) Then Exit Function
    FolderPath = pickedFolder & "\Reports - " & RuleProfile.Name & " - " & SetAssets.Name & " (" & Format$(Now(), "yyyy-mm-dd-hh-mm-ss") & ")"
    If Not fso.FolderExists(FolderPath) Then
        MkDir FolderPath
    End If
    pickReportFolder = FolderPath
End Function

Private Sub toggleAreaVisibility(ByVal visRange As SvnMultiRange, ByVal newHidden As Boolean, Optional ByVal CopyFriendly As Boolean = True)
    If Not visRange Is Nothing Then
        Dim currentArea As Range
        For Each currentArea In visRange.Areas
            If CopyFriendly Then
                Dim targetSheet As Worksheet
                Set targetSheet = CopyFriendlyGet(currentArea.Worksheet.Name, False) ' Don't mess with the hide ranges on sheets that don't exist yet - they'll get fixed later
                If Not targetSheet Is Nothing Then
                    targetSheet.Range(currentArea.Address).EntireRow.Hidden = newHidden
                End If
            Else
                currentArea.Worksheet.Range(currentArea.Address).EntireRow.Hidden = newHidden
            End If
        Next
    End If
End Sub

Private Function getSwapSheets(ByVal givenSheets As Collection) As Collection ' Also refreshes values on dupes
    Dim outSheets As Collection
    Set outSheets = New Collection
    Dim formulaSheet As Variant
    For Each formulaSheet In givenSheets
        Dim valueSheet As Worksheet
        Set valueSheet = CopyFriendlyGet(formulaSheet.Name)
        If valueSheet.Name <> formulaSheet.Name Then valueSheet.Range(formulaSheet.UsedRange.Address).Value2 = formulaSheet.UsedRange.Value2
        If collectionGetOrNothing(outSheets, valueSheet.Name) Is Nothing Then outSheets.Add valueSheet, valueSheet.Name
    Next
    Set getSwapSheets = outSheets
End Function

Private Function CopyFriendlyGet(ByVal originalSheetName As String, Optional ByVal allowNewCopies As Variant = True) As Worksheet ' Gets copy if relevant, original if not, makes copy if neccesary
    Dim CopyName As String
    CopyName = originalSheetName
    For comboIndex = 1 To currentCombo.count
        If Not collectionGetOrNothing(GetRelevantSwap(comboIndex).DupeSheets, originalSheetName) Is Nothing And GetSwapIndex(comboIndex) > 1 Then CopyName = CopyName & "-" & currentCombo(comboIndex)
    Next
    
    If CopyName = originalSheetName Then         ' Original Sheet
        Set CopyFriendlyGet = ThisWorkbook.Sheets(originalSheetName)
    ElseIf Not collectionGetOrNothing(copiedSheets, CopyName) Is Nothing Then ' Already Copied
        Set CopyFriendlyGet = copiedSheets(CopyName)
    ElseIf allowNewCopies Then                   ' Otherwise leave as nothing
        ' Find 'Before' Sheet
        Dim beforeSheet As Worksheet
        Dim searchSheet As Long
        Dim copyIndex As Long
        copyIndex = DupeIndexes(CopyName)        ' Fails if not found
        For searchSheet = 1 To ThisWorkbook.Sheets.count
            Dim searchIndex As Variant
            searchIndex = collectionGetOrEmpty(DupeIndexes, ThisWorkbook.Sheets(searchSheet).Name)
            If Not IsEmpty(searchIndex) Then
                If copyIndex < searchIndex Then
                    Set beforeSheet = ThisWorkbook.Sheets(searchSheet - 1) ' Invariant: searchindex should never be 1
                    Exit For
                End If
            End If
        Next
        
        If beforeSheet Is Nothing Then Set beforeSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count) ' Add it last
    
        ' Make Copy
        DoEvents
        FlnUI.SetStatus logOutput, inLine4:="Duplicating sheet '" & originalSheetName & "'..."
        ThisWorkbook.Sheets(originalSheetName).Copy After:=beforeSheet ' Copy
        ActiveSheet.Name = CopyName              ' Rename
        ThisWorkbook.Sheets(CopyName).Unprotect vbNullString ' Allow pasting over
        copiedSheets.Add ThisWorkbook.Sheets(CopyName), CopyName ' Add sheet before select or DoEvents to avoid missed sheets
        Set CopyFriendlyGet = ThisWorkbook.Sheets(CopyName)
        worksheetSetProperty ThisWorkbook, propertyFromInstantiate(DUPE_PROPERTY, True, msoPropertyTypeString, CopyFriendlyGet)
        DoEvents
    End If
End Function

Public Sub GetDupeIndexes()
    Set DupeIndexes = New Collection
    Dim workingCollection As Collection
    Set workingCollection = New Collection
    Dim currentSheet As Worksheet
    For Each currentSheet In collectionGetOrderedSheets(RuleProfile.GetSheets, ThisWorkbook)
        workingCollection.Add propertyFromInstantiate(currentSheet.Name, currentSheet.Name, msoPropertyTypeString), currentSheet.Name
    Next
    
    Dim SwapCollection As Collection
    Dim currentKey As Long
    Dim propnum As Long
    For Each currentSwap In Swaps
        ' Set temporary Collection
        Set SwapCollection = New Collection
        For propnum = 1 To workingCollection.count
            SwapCollection.Add workingCollection(propnum), workingCollection(propnum).value
        Next
        
        For currentKey = 2 To currentSwap.MaxSwapSet.Rows.count
            For propnum = 1 To workingCollection.count
                If Not collectionGetOrNothing(currentSwap.DupeSheets, workingCollection(propnum).Name) Is Nothing Then ' Matching Sheet
                    If currentSwap.doEndOrder Or propnum = workingCollection.count Then
                        SwapCollection.Add propertyFromInstantiate(workingCollection(propnum).Name, _
                                                                   workingCollection(propnum).value & "-" & CStr(currentSwap.MaxSwapSet.Rows(currentKey).Value2), msoPropertyTypeString)
                    Else
                        SwapCollection.Add propertyFromInstantiate(workingCollection(propnum).Name, _
                                                                   workingCollection(propnum).value & "-" & CStr(currentSwap.MaxSwapSet.Rows(currentKey).Value2), msoPropertyTypeString), _
        before:=workingCollection(propnum).value
                    End If
                End If
            Next
        Next
        
        ' Update Real collection
        Set workingCollection = SwapCollection
    Next
    
    Dim i As Long
    For i = 1 To workingCollection.count
        DupeIndexes.Add i, workingCollection(i).value ' Index keyed by sheet name
    Next
End Sub

Private Function GetSwapIndex(ByVal inComboIndex As Variant) As Long
    Dim MaxSet As Range
    Set MaxSet = GetRelevantSwap(inComboIndex).MaxSwapSet
    Dim currentRow As Long
    For currentRow = 1 To MaxSet.Rows.count
        If MaxSet.Rows(currentRow).Value2 = currentCombo(inComboIndex) Then GetSwapIndex = currentRow
    Next
End Function

Private Function GetRelevantSwap(ByVal comboIndex As Long) As FlnSwap
    Dim swapIndex As Long
    Dim maxIndex As Long
    maxIndex = comboIndex
    For swapIndex = 1 To maxIndex
        If Swaps(swapIndex).swapRange Is Nothing Then maxIndex = comboIndex + 1 ' Offensive
    Next
    Set GetRelevantSwap = Swaps(comboIndex)
End Function

Public Function AddPhotos(Optional ByVal CopyFriendly As Boolean = True) As SvnResult
    Set AddPhotos = resultFromInstantiate(svnSuccess)
    
    Dim CurrentPhotos As Range
    Dim currentDestSheet As Worksheet
    Dim PhotoDest As Range
    Dim photoSplit() As String
    Dim photoArea As Long
    For Each currentPhotoFill In photofills
        ' Get Destination Worksheet
        If CopyFriendly Then Set currentDestSheet = CopyFriendlyGet(currentPhotoFill.Dest.Worksheet.Name, False)
        If Not CopyFriendly Then Set currentDestSheet = currentPhotoFill.Dest.Worksheet
        If collectionGetOrNothing(AssetSheets, currentDestSheet.Name) Is Nothing Then GoTo NextPhoto
        For Each currentPhotoShape In addedShapes
            If currentPhotoShape.title = currentPhotoFill.Name & currentDestSheet.Name Then GoTo NextPhoto ' Skip Adding photos if they already exist
        Next
        
        For photoArea = 1 To currentPhotoFill.Source.Areas.count
                    
            Set CurrentPhotos = currentPhotoFill.Source.Areas(photoArea)
            If IsError(CurrentPhotos) Then
                If errCancelResult("Photo Source is an Error", currentPhotoFill.Name & " Area " & CStr(photoArea) & "(" & CurrentPhotos.Address & ", " & CurrentPhotos.Text & ")", AddPhotos) Then Exit Function
            End If
            
            Set PhotoDest = currentDestSheet.Range(currentPhotoFill.Dest.Areas(photoArea).Address)
                        
            photoSplit = Split(CurrentPhotos.value, ",")
            If UBound(photoSplit) - LBound(photoSplit) > PhotoDest.Rows.count - 1 Then
                If errCancelResult("Not Enough Image Cells", CStr(UBound(photoSplit) - LBound(photoSplit) + 1) & " Images", AddPhotos) Then Exit Function
            End If
            
            Dim currentRowIndex As Long
            Dim PhotoIndex As Long
            For currentRowIndex = 1 To PhotoDest.Rows.count
                PhotoIndex = currentRowIndex - 1
                If PhotoIndex <= UBound(photoSplit) Then
                    Dim FileName As String
                    FileName = replace(photoPath, "%WORKBOOKPATH%", ThisWorkbook.Path) & "\" & photoSplit(PhotoIndex)
                    If Dir(FileName & ".jpg") <> vbNullString Or Dir(FileName & ".png") <> vbNullString Then
                        addedShapes.Add rangeInsertPicture(IIf(Dir(FileName & ".jpg") <> vbNullString, FileName & ".jpg", FileName & ".png"), PhotoDest, PHOTO_TITLE)
                    Else
                        If errCancelResult("Image Not Found", CStr(FileName), AddPhotos) Then Exit Function
                    End If
                End If
            Next
        Next
NextPhoto:
    Next
End Function

Public Function GetSwapCombos() As SvnResult
    If Not pDebug Then On Error GoTo SubHandler Else On Error GoTo 0
    If Not pDebug Then Application.EnableCancelKey = xlErrorHandler Else Application.EnableCancelKey = xlInterrupt

    Set GetSwapCombos = resultFromInstantiate(svnSuccess)
    
    Set SwapCombos = New Collection
    Dim newCombos As Collection
        
    Dim currentCell As Long
    Dim currentLeft As Collection
    Dim currentRight As Long
    Dim leftelement As Variant

    ' Get cartesian product of swapcombos and each swapset
    Dim swapNum As Long
    For swapNum = 1 To Swaps.count
        Set currentSwap = Swaps(swapNum)
        If SwapCombos.count = 0 And Not currentSwap.swapRange Is Nothing Then ' Initial fill if empty - Top level must not be dependent on lower level, must have at least 1 element. fix by counting in nothings
            For currentCell = currentSwap.swapRange.Rows.count To 1 Step -1
                SwapCombos.Add New Collection
                SwapCombos(SwapCombos.count).Add currentSwap.swapRange.Rows(currentCell).Value2
            Next
        Else                                     ' Dependent on top level
            Set newCombos = New Collection
            For Each currentLeft In SwapCombos
                ' Apply cells and reload
                For comboIndex = 1 To currentLeft.count
                    Swaps(comboIndex).SwapCell.Value2 = currentLeft(comboIndex)
                Next
                Application.EnableCancelKey = xlDisabled
                On Error GoTo 0                  ' Don't Catch UDF Cell Errors
                Application.Calculate
                If Not pDebug Then On Error GoTo SubHandler Else On Error GoTo 0
                If Not pDebug Then Application.EnableCancelKey = xlErrorHandler Else Application.EnableCancelKey = xlInterrupt
                DoEvents
                    
                If currentSwap.swapRange Is Nothing Then
                    currentLeft.Add Empty
                    newCombos.Add currentLeft
                Else
                    For currentRight = currentSwap.swapRange.Rows.count To 1 Step -1
                        newCombos.Add New Collection
                        For Each leftelement In currentLeft
                            newCombos(newCombos.count).Add leftelement
                        Next
                        newCombos(newCombos.count).Add currentSwap.swapRange.Rows(currentRight).Value2
                    Next
                End If
            Next
            Set SwapCombos = newCombos
        End If
    Next
        
    ' Remove All Nothings - Key by Swaps here to remove need for getrelevantswap indexing?
    Dim currentElement As Long
    For Each currentCombo In SwapCombos
        currentElement = 1
        Do While currentElement <= currentCombo.count
            If IsEmpty(currentCombo(currentElement)) Then
                currentCombo.Remove currentElement
            Else
                currentElement = currentElement + 1
            End If
        Loop
    Next
    
    GetSwapCombos.Instantiate svnSuccess
Finish:
    Exit Function
SubHandler:
    GetSwapCombos.Instantiate SvnSubErr, Err.Number, Err.Description
    Resume Finish
End Function

Public Function applyAutofit(Optional ByVal CopyFriendly As Boolean = True) As SvnResult
    Set applyAutofit = resultFromInstantiate(svnSuccess)
    Dim currentAutofit As FlnAutoFit
    Dim currentRangeSheet As Worksheet
    Dim copyFriendlySheet As Worksheet
    Dim currentArea As Range
    Dim adjustedArea As Range
    Dim currentRow As Range
    Dim currentColumn As Range
    Dim oldWidth As Double
    Dim mergedWidth As Double
    Dim oldMerge As Range
    
    For Each currentAutofit In autofits
        For Each currentRangeSheet In currentAutofit.MultiRange.Worksheets
        
            If CopyFriendly Then Set copyFriendlySheet = CopyFriendlyGet(currentRangeSheet.Name, False)
            If Not CopyFriendly Then Set copyFriendlySheet = currentRangeSheet
            If Not collectionGetOrNothing(AssetSheets, copyFriendlySheet.Name) Is Nothing Then
                
                For Each currentArea In currentAutofit.MultiRange.AreasBySheet(currentRangeSheet)
                
                    Set adjustedArea = copyFriendlySheet.Range(currentArea.Address)
                
                    For Each currentRow In adjustedArea.Rows
                        If currentRow.MergeCells Then
                            mergedWidth = 0
                            If currentRow.Columns.count > 1 Then
                                For Each currentColumn In currentRow.Columns
                                    mergedWidth = mergedWidth + currentColumn.ColumnWidth
                                Next
                                Set oldMerge = currentRow
                            Else
                                For Each currentColumn In currentRow.MergeArea.Columns
                                    mergedWidth = mergedWidth + currentColumn.ColumnWidth
                                Next
                                Set oldMerge = currentRow.MergeArea
                            End If
    
                            currentRow.MergeCells = False
                            oldWidth = currentRow.Columns(1).ColumnWidth
                            currentRow.Columns(1).ColumnWidth = mergedWidth
                            currentRow.Columns(1).EntireRow.AutoFit
                            oldMerge.Merge True
                            currentRow.Columns(1).ColumnWidth = oldWidth
                        Else
                            currentRow.AutoFit
                        End If
                    Next
                Next
            End If
        Next
    Next
    applyAutofit.Instantiate svnSuccess
End Function

Public Function applyAutohide(Optional ByVal CopyFriendly As Boolean = True) As SvnResult
    Set applyAutohide = resultFromInstantiate(svnSuccess)
    Dim currentAutohide As FlnAutoHide
    Dim currentRangeSheet As Worksheet
    Dim copyFriendlyRangeSheet As Worksheet
    Dim adjustedArea As Range
    Dim currentArea As Range
    Dim currentRow As Range
    Dim currentColumn As Range
    
    
    For Each currentAutohide In autohides
        For Each currentRangeSheet In currentAutohide.MultiRange.Worksheets
            If CopyFriendly Then Set copyFriendlyRangeSheet = CopyFriendlyGet(currentRangeSheet.Name, False)
            If Not CopyFriendly Then Set copyFriendlyRangeSheet = currentRangeSheet
            If Not collectionGetOrNothing(AssetSheets, copyFriendlyRangeSheet.Name) Is Nothing Then
                For Each currentArea In currentAutohide.MultiRange.AreasBySheet(copyFriendlyRangeSheet)
                    Set adjustedArea = copyFriendlyRangeSheet.Range(currentArea.Address)
                    For Each currentRow In adjustedArea.Rows
                        If currentRow.Value2 = True Then
                            currentRow.EntireRow.Hidden = True
                        ElseIf currentRow.Value2 = False Then
                            currentRow.EntireRow.Hidden = False
                        End If
                    Next
                Next
            End If
        Next
    Next
    applyAutohide.Instantiate svnSuccess
    
End Function

Public Function applyPageGroups(Optional ByVal CopyFriendly As Boolean = True) As SvnResult
    Set applyPageGroups = resultFromInstantiate(svnSuccess)
    
    Dim currentPageGroup As FlnPageGroup
    Dim currentRangeSheet As Worksheet
    Dim copyFriendlyRangeSheet As Worksheet
    Dim adjustedArea As Range
    Dim currentArea As Range
    Dim mergedArea As Range
    Dim currentRow As Range
    
    ' XXX - Remove all page breaks on page group pages
    For Each currentPageGroup In pageGroups
        For Each currentRangeSheet In currentPageGroup.MultiRange.Worksheets
            currentRangeSheet.ResetAllPageBreaks
        Next
    Next
    
    For Each currentPageGroup In pageGroups
        For Each currentRangeSheet In currentPageGroup.MultiRange.Worksheets
        
            If CopyFriendly Then Set copyFriendlyRangeSheet = CopyFriendlyGet(currentRangeSheet.Name, False)
            If Not CopyFriendly Then Set copyFriendlyRangeSheet = currentRangeSheet
            
            If Not collectionGetOrNothing(AssetSheets, copyFriendlyRangeSheet.Name) Is Nothing Then
            
                ' Process Page Groups
                For Each currentArea In currentPageGroup.MultiRange.AreasBySheet(currentRangeSheet)
                    ' Handle Seperate Areas Seperately
                    
                    Set adjustedArea = copyFriendlyRangeSheet.Range(currentArea.Address)
                    
                    If currentArea.EntireRow.Hidden = False Then
                        If currentArea.MergeCells Then Set mergedArea = currentArea.MergeArea
                        If Not currentArea.MergeCells Or IsNull(currentArea.MergeCells) Then Set mergedArea = currentArea
                        For Each currentRow In mergedArea.Rows
                            If currentRow.EntireRow.PageBreak = xlPageBreakAutomatic Then
                                If mergedArea.Rows(1).EntireRow.PageBreak <> xlPageBreakManual Then
                                    mergedArea.Rows(1).EntireRow.PageBreak = xlPageBreakManual
                                    addedBreaks.Add mergedArea.Rows(1).EntireRow
                                End If
                            End If
                        Next
                    End If
                Next
                
            End If
        Next
    Next
    
    applyPageGroups.Instantiate svnSuccess
End Function

Private Function errCancel(ByVal errorText As String, ByVal identifierText As String) As Boolean
    ''' Display Error Message and Prompt '''
    errCancel = IIf((MsgBox(errorText & " on ID: " & CStr(AssetCell.Value2) & " - " & identifierText & vbCrLf & "Continue to the next asset anyway?", vbOKCancel) <> vbOK), True, False)
End Function

Private Function errCancelResult(ByVal errorText As String, ByVal identifierText As String, ByVal inSvnResult As SvnResult) As Boolean
    If errCancel(errorText, identifierText) Then inSvnResult.Instantiate svnCancel Else inSvnResult.Instantiate svnSkip
    errCancelResult = True
End Function

Private Function SafeDir(ByVal Path As String, ByVal inAttr As VbFileAttribute) As String
    On Error Resume Next
    SafeDir = Dir(Path, inAttr)
    On Error GoTo 0
End Function
'' ** Signature Mirrored in FlnFlorin ** ''
Public Function Instantiate(ByVal inProfile As FlnProfile, ByVal inAssets As FlnAssetList, ByVal inDebug As Boolean) As FlnGenerator
    ''' Initialize Properties '''
    Set RuleProfile = inProfile
    Set SetAssets = inAssets
    pDebug = inDebug
    
    ' Expand Profile
    If Not RuleProfile Is Nothing Then
        Set OutputTables = RuleProfile.OutputTables
        Set AssetCell = RuleProfile.AssetCell
        Set FilenameCell = RuleProfile.FilenameCell
        Set Swaps = RuleProfile.Swaps
        Set photofills = RuleProfile.photofills
        Set autofits = RuleProfile.autofits
        Set autohides = RuleProfile.autohides
        Set pageGroups = RuleProfile.pageGroups
        photoPath = RuleProfile.photoPath
    Else
        Set OutputTables = New Collection
        Set Swaps = New Collection
        Set photofills = New Collection
        Set autofits = New Collection
        Set autohides = New Collection
        Set pageGroups = New Collection
        photoPath = vbNullString
    End If

    Set Instantiate = Me
End Function


