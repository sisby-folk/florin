Attribute VB_Name = "FlnMacros"
'@Folder("Florin")
'@IgnoreModule ProcedureNotUsed
Option Explicit


Public Sub openFlorin()
    ''' Macro: Open UserForm '''
    Application.EnableCancelKey = xlInterrupt
    FlnUI.Show
    Application.EnableCancelKey = xlInterrupt
End Sub

Public Sub shapeAlignmentTest()
    Dim rangeString As String
    rangeString = stringFromRefEdit(ThisWorkbook, "Enter Range to Test over", "Test Shape Alignment", Dimensions:=dimensionsFromRowsColsRepeatable(False, -1, -1))
    If rangeString = vbNullString Then ' Valid
        MsgBox "Reference entered was invalid for parameter", vbExclamation & vbOKOnly, "Invalid Reference"
        Exit Sub
    End If
    Dim testRange As Range
    Set testRange = rangeOrNothing(rangeString, ActiveWorkbook)
    Dim area As Range
    Dim newShape As Shape
    For Each area In testRange.Areas
        If (area.MergeCells) Then
            Set newShape = area.Worksheet.Shapes.AddShape(msoShape10pointStar, area.Left + area.MergeArea.Width * 0.05, area.Top + area.MergeArea.Height * 0.05, area.MergeArea.Width * 0.9, area.MergeArea.Height * 0.9)
        Else
            Set newShape = area.Worksheet.Shapes.AddShape(msoShape10pointStar, area.Left + area.Width * 0.05, area.Top + area.Height * 0.05, area.Width * 0.9, area.Height * 0.9)
        End If
        newShape.title = "TestShape"
    Next
End Sub

Public Sub manualAddPhotos()
    Dim Profiles As Collection
    Set Profiles = FlnUI.GetProfilesPublic()
    
    Dim photofills As Collection
    Set photofills = Profiles(1).photofills
    
    Dim currentPhotoFill As FlnPhotoFill
    Dim CurrentPhotos As Range
    Dim currentDestSheet As Worksheet
    Dim PhotoDest As Range
    Dim photoSplit() As String
    Dim photoArea As Long
    For Each currentPhotoFill In photofills
        ' Get Destination Worksheet
        Set currentDestSheet = currentPhotoFill.Dest.Worksheet
        
        For photoArea = 1 To currentPhotoFill.Source.Areas.count
                    
            Set CurrentPhotos = currentPhotoFill.Source.Areas(photoArea)
            If IsError(CurrentPhotos) Then
                MsgBox "Photo Source is an Error"
                Exit Sub
            End If
            
            Set PhotoDest = currentDestSheet.Range(currentPhotoFill.Dest.Areas(photoArea).Address)
                        
            photoSplit = Split(CurrentPhotos.value, ",")
            If UBound(photoSplit) - LBound(photoSplit) > PhotoDest.Rows.count - 1 Then
                MsgBox "Not Enough Image Cells"
                Exit Sub
            End If
            
            Dim currentRowIndex As Long
            Dim PhotoIndex As Long
            For currentRowIndex = 1 To PhotoDest.Rows.count
                PhotoIndex = currentRowIndex - 1
                If PhotoIndex <= UBound(photoSplit) Then
                    Dim FileName As String
                    FileName = Dir(replace(Profiles(1).photoPath, "%WORKBOOKPATH%", ThisWorkbook.Path), vbDirectory) & "\" & photoSplit(PhotoIndex)
                    If Dir(FileName & ".jpg") <> vbNullString Or Dir(FileName & ".png") <> vbNullString Then
                        PhotoDest.Rows(currentRowIndex).MergeArea.EntireRow.Hidden = False
                        Dim newPicShape As Shape
                        Set newPicShape = PhotoDest.Worksheet.Shapes.AddPicture(IIf(Dir(FileName & ".jpg") <> vbNullString, FileName & ".jpg", FileName & ".png"), msoTrue, msoFalse, PhotoDest.Rows(currentRowIndex).Left, PhotoDest.Rows(currentRowIndex).Top + 4, -1, -1)
                        newPicShape.title = currentPhotoFill.Name & PhotoDest.Worksheet.Name
                        newPicShape.LockAspectRatio = msoTrue
                        newPicShape.Width = PhotoDest.MergeArea.Width * 0.8
                        newPicShape.Left = newPicShape.Left + PhotoDest.MergeArea.Width * 0.1
                        If newPicShape.Height > PhotoDest.MergeArea.Height * 0.8 Then newPicShape.Height = PhotoDest.MergeArea.Height * 0.8
                    Else
                        MsgBox "Image Not Found"
                        Exit Sub
                    End If
                End If
            Next
        Next
    Next
End Sub

Public Sub spaceSet()
    Dim vertSpacing As Long
    Dim setRange As Range
    Dim destinationCell As Range
    vertSpacing = ActiveCell.MergeArea.Rows.count
    Set destinationCell = ActiveCell.Cells(1, 1)
    Dim rangeString As String
    rangeString = stringFromRefEdit(ThisWorkbook, "Enter Range to Space by " & CStr(vertSpacing), "Space Set", Dimensions:=dimensionsFromRowsCols(-1, -1), Unionize:=True)
    If rangeString = vbNullString Then
        MsgBox "Reference entered was invalid for parameter", vbExclamation & vbOKOnly, "Invalid Reference"
        Exit Sub
    End If
    
    Set setRange = rangeOrNothing(rangeString, ActiveWorkbook)

    Dim currentCell As Range
    Dim firstCell As Boolean
    firstCell = True
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For Each currentCell In setRange.Rows
        destinationCell.Value2 = currentCell.Value2
        If Not firstCell Then destinationCell.Worksheet.Range(destinationCell, destinationCell.Rows(vertSpacing)).Merge
        If firstCell Then firstCell = False
        Set destinationCell = destinationCell.Rows(vertSpacing + 1)
    Next
    On Error Resume Next
    Application.Calculate
    On Error GoTo 0
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Sub mergeAndRecolor()

    Dim ColorRangeString As String
    ColorRangeString = stringFromRefEdit(ActiveWorkbook, vbNullString, "Rows To Color...", Dimensions:=dimensionsFromRowsCols(-1, -1), Unionize:=True)
    If ColorRangeString = vbNullString Then
        MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
        Exit Sub
    End If
    Dim ColorRange As Range
    Set ColorRange = rangeOrNothing(ColorRangeString, ActiveWorkbook)

    Dim SwapColorRangeString As String
    SwapColorRangeString = stringFromRefEdit(ActiveWorkbook, ColorRangeString, "Merge/Color By...", Dimensions:=dimensionsFromRowsCols(ColorRange.Rows.count, 1), Unionize:=True)
    If SwapColorRangeString = vbNullString Then
        MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
        Exit Sub
    End If
    Dim SwapColorRange As Range
    Set SwapColorRange = rangeOrNothing(SwapColorRangeString, ActiveWorkbook)
    
    Dim MergeRangeString As String
    MergeRangeString = stringFromRefEdit(ActiveWorkbook, MergeRangeString, "Columns To Merge", Dimensions:=dimensionsFromRowsColsRepeatable(False, ColorRange.Rows.count, -1))
    If MergeRangeString = vbNullString Then
        MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
        Exit Sub
    End If
    Dim MergeRange As Range
    Set MergeRange = rangeOrNothing(MergeRangeString, ActiveWorkbook)
    
    Dim DoBlue As Boolean
    Dim rowStreak As Long
    Dim currentArea As Range
    Dim currentColumn As Range
    Dim currentRow As Long
    Dim includeCurrent As Boolean
    For currentRow = 1 To ColorRange.Rows.count
        If Not IsEmpty(SwapColorRange.Rows(currentRow)) Then DoBlue = Not DoBlue
        
        For Each currentArea In ColorRange.Areas
            currentArea.Rows(currentRow).Interior.Color = IIf(DoBlue, RGB(191, 191, 191), RGB(230, 230, 241))
        Next
        
        If rowStreak > 0 And Not IsEmpty(SwapColorRange.Rows(currentRow)) Or currentRow = ColorRange.Rows.count Then ' Merge if appropriate
            Application.DisplayAlerts = False
            includeCurrent = currentRow = ColorRange.Rows.count And IsEmpty(IsEmpty(SwapColorRange.Rows(currentRow)))
            For Each currentArea In MergeRange.Areas
                For Each currentColumn In currentArea.Columns
                    currentColumn.Rows(currentRow - rowStreak).Resize(rowStreak + IIf(includeCurrent, 1, 0)).Merge
                Next
            Next
            rowStreak = 0
            Application.DisplayAlerts = True
        End If
        
        rowStreak = rowStreak + 1
    Next
End Sub

