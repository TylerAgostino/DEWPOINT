Function loadFavorites

	'First open the favorites file
	Set favesFile = CreateObject("Scripting.FileSystemObject").OpenTextFile("include/favorites.txt",1)
	'And load it's contents into a variable
	strFileText = favesFile.ReadAll()
	favesFile.Close
	'Then replace the semicolons with the table cell tags, and line breaks with the table row tags
	strFileTextFixed = Replace(strFileText, ";", "'></td><td>")
	strFileTextFixed = Replace(strFileTextFixed, vbCrLf, "</td></tr><tr><td align='right'><input type='radio' name='weatherstation'  value='")
	
	loadFavorites="<tr><td align='right'><input type='radio' name='weatherstation' value='"+strFileTextFixed+"</td></tr>"
	'outputSpan.innerhtml="test"
End Function
