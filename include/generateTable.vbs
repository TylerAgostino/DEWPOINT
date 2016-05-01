Function generateTable
	'First open the favorites file
	Set tableTop = CreateObject("Scripting.FileSystemObject").OpenTextFile("include/tableTop.txt",1)
	'And load it's contents into a variable
	topHTML = tableTop.ReadAll()
	tableTop.Close
	
	favsHTML = loadFavorites
	
	'First open the favorites file
	Set tableBottom = CreateObject("Scripting.FileSystemObject").OpenTextFile("include/tableBottom.txt",1)
	'And load it's contents into a variable
	bottomHTML = tableBottom.ReadAll()
	tableBottom.Close
	
	outputHTML = topHTML+favsHTML+bottomHTML
	
	set output = document.getelementbyid("outputSpan")
	output.innerHTML = outputHTML
	
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile("C:\Shortcuts\output.txt",True)
	objFile.Write outputHTML
	objFile.Close
	
end Function