VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FlnProfileEditor 
   Caption         =   "Florin Profile Editor"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   OleObjectBlob   =   "FlnProfileEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FlnProfileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Florin")
Option Explicit

Private Profiles As Collection
Private pDirty As Boolean
Private pReturnIndex As Long

Public Property Get ReturnIndex() As Long
    ReturnIndex = pReturnIndex
End Property

Public Function isDirty() As Boolean
    isDirty = pDirty
End Function

Private Sub DirtyRules()
    pDirty = True
    LoadForm
End Sub

Private Sub reOrder(ByVal changeBox As MSForms.ListBox, ByVal changeCol As Collection, ByVal changeIndex As Long, ByVal directionUp As Boolean)
    If collectionMoveDirection(changeCol, changeIndex, directionUp) Then
        changeBox.ListIndex = changeIndex - 1 + IIf(directionUp, -1, 1)
        DirtyRules
    End If
End Sub

Private Sub LoadForm()
    Dim profileIndex As Long
    profileIndex = ListProfiles.ListIndex
    
    ListProfiles.Clear
    
    Dim CurrentProfile As FlnProfile
    For Each CurrentProfile In Profiles
        ListProfiles.AddItem
        ListProfiles.list(ListProfiles.ListCount - 1, 0) = CurrentProfile.Name
        ' array multiply for each flag to ensure they always appear in the same order
        ListProfiles.list(ListProfiles.ListCount - 1, 1) = arrayGetJoin(arrayBoolMult(arrayFrom1DArgs("E", "P", "R", "S", "O"), arrayFrom1DArgs(CurrentProfile.flagDoesExport, CurrentProfile.flagHasPhotos, CurrentProfile.flagRenamesPhotos, CurrentProfile.flagHasSwaps, CurrentProfile.flagHasOutputs), ""), "")
    Next
    If profileIndex < ListProfiles.ListCount Then ListProfiles.ListIndex = profileIndex
    
End Sub

Public Sub Initialize(ByVal inProfiles As Collection)
    pDirty = False
    Set Profiles = inProfiles
    LoadForm
End Sub

Private Sub ButtonNewProfile_Click()
    Dim newProfName As String
    newProfName = Left$(replace(InputBox("Enter New Profile Name:", "Edit Profiles"), " ", vbNullString), 100)
    If newProfName <> vbNullString And collectionGetOrNothing(Profiles, newProfName) Is Nothing Then
        Profiles.Add profileFromInstantiate(newProfName, New Collection, New Collection, New Collection, New Collection, New Collection, New Collection, New Collection, New Collection, Nothing, Nothing, vbNullString, Nothing), newProfName
        DirtyRules
    End If
End Sub

Private Sub ButtonProfileSelect_Click()
    pReturnIndex = ListProfiles.ListIndex
    Me.Hide
End Sub

Private Sub ButtonRemoveProfile_Click()
    If ListProfiles.ListIndex <> -1 And MsgBox("This will erase ALL profile configuration - Rules, Swaps, Asset Lists, AND Outputs, Are you sure?", vbYesNo, "Delete Profile") = vbYes Then
        Profiles.Remove ListProfiles.ListIndex + 1
        DirtyRules
    End If
End Sub

Private Sub ButtonEditProfile_Click()
    If ListProfiles.ListIndex <> -1 Then
        Dim newProfileName As String
        newProfileName = Left$(replace(InputBox("Enter New Profile Name:", "Edit Profiles"), " ", vbNullString), 100)
        If newProfileName <> vbNullString And collectionGetOrNothing(Profiles, newProfileName) Is Nothing Then
            Profiles.Add Profiles(ListProfiles.ListIndex + 1), newProfileName, After:=ListProfiles.ListIndex + 1 'Duplicate Reference with new key
            Profiles.Remove ListProfiles.ListIndex + 1
            Profiles(newProfileName).Name = newProfileName
            DirtyRules
        End If
    End If
End Sub

Private Sub ButtonProfileDown_Click()
    If ListProfiles.ListIndex <> -1 Then reOrder ListProfiles, Profiles, ListProfiles.ListIndex + 1, False
End Sub

Private Sub ButtonProfileUp_Click()
    If ListProfiles.ListIndex <> -1 Then reOrder ListProfiles, Profiles, ListProfiles.ListIndex + 1, True
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.Hide
    Cancel = True
End Sub

