VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FlnCaseEditor 
   Caption         =   "Florin Case Editor"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   OleObjectBlob   =   "FlnCaseEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FlnCaseEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("Florin")
Option Explicit

Private thisCase As FlnCase
Private pDirty As Boolean

Private Enum RefProperties
    Condition = 0
    HideRef = 1
End Enum

Public Property Get isDirty() As Boolean
    isDirty = pDirty
End Property

Public Property Get FormCase() As FlnCase
    Set FormCase = thisCase
End Property

Private Sub LoadSheets()
    ListSheets.Clear
    Dim currentSheet As Worksheet
    For Each currentSheet In thisCase.Sheets
        ListSheets.AddItem currentSheet.Name
    Next
End Sub

Private Sub LoadForm()
    ButtonRefEditCondition.Caption = thisCase.Condition
    ButtonRefEditHide.Caption = thisCase.HideRef.ToString
    LoadSheets
End Sub

Private Sub Dirty()
    pDirty = True
    LoadForm
End Sub

Private Sub RefEdit()
    With New SvnFormulaEditor
        Dim oldProp As String
        oldProp = thisCase.Condition
        If oldProp = vbNullString Then
            oldProp = thisCase.HideRef.ToString
        End If
        .Initialize oldProp, "Edit Condition"
        Me.Hide
        .Show vbModal
        If Not IsEmpty(varOrEmpty(.Formula)) Then
            thisCase.Condition = IIf(Left(.Formula, 1) = "=", vbNullString, "=") & IIf(.Formula = vbNullString, Empty, .Formula)
            Dirty
        Else
            MsgBox "Reference entered was invalid for property", vbExclamation & vbOKOnly, "Invalid Reference"
        End If
    End With
    Me.Show
End Sub

Private Sub AddSheet()
    Dim sheetsToPick As Collection
    Set sheetsToPick = New Collection
    Dim currentSheet As Worksheet
    For Each currentSheet In ThisWorkbook.Sheets
        If collectionGetOrNothing(thisCase.Sheets, currentSheet.Name) Is Nothing Then
            sheetsToPick.Add currentSheet.Name
        End If
    Next
    If sheetsToPick.count > 0 Then
        With New SvnPicker
            .Initialize "Pick Sheet to Add", sheetsToPick
            .Show
            If .Success Then
                thisCase.Sheets.Add ThisWorkbook.Sheets(sheetsToPick(.Index)), sheetsToPick(.Index)
                Dirty
            End If
        End With
    End If
End Sub

Private Sub RemoveSheet()
    If ListSheets.ListIndex <> -1 Then
        thisCase.Sheets.Remove ListSheets.ListIndex + 1
        Dirty
    End If
End Sub

Public Sub Initialize(ByVal inCase As FlnCase, Optional ByVal ruleName As String = vbNullString, Optional ByVal caseNum As Long = -1)
    pDirty = False
    Set thisCase = inCase
    If ruleName = vbNullString Then
        LabelHeader.Caption = "Ranges"
    Else
        LabelHeader.Caption = ruleName & ": Case " & CStr(caseNum)
    End If
    LoadForm
End Sub

Private Sub ButtonAddSheet_Click()
    AddSheet
End Sub

Private Sub ButtonCancel_Click()
    pDirty = False
    Me.Hide
End Sub

Private Sub ButtonRefEditCondition_Click()
    RefEdit
End Sub

Private Sub ButtonRefEditHide_Click()
    If thisCase.HideRef.Edit("Hide Range - " & LabelHeader.Caption, dimensionsFromRowsColsRepeatable(True, -1, -1)) Then Dirty
End Sub

Private Sub ButtonRemoveSheet_Click()
    RemoveSheet
End Sub

Private Sub ButtonSave_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ButtonCancel_Click
    Cancel = True
End Sub

