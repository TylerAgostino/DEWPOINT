Sub ChooseSaveFolder
    strStartDir = "c:\"
    userInput.saveLocation.value = PickFolder(strStartDir)
End Sub 

Function PickFolder(strStartDir)
	Dim SA, F
	Set SA = CreateObject("Shell.Application")
	Set F = SA.BrowseForFolder(0, "Choose a folder", 0)
	If (Not F Is Nothing) Then
	  PickFolder = F.Items.Item.path
	End If
	Set F = Nothing
	Set SA = Nothing
End Function 