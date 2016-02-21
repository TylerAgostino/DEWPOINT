'************************************************************************************
'*																					*
'*							Copyright 2016 Tyler Agostino 							*
'*																					*
'*		This program is free software: you can redistribute it and/or modify		*
'*		it under the terms of the GNU General Public License as published by		*
'*		the Free Software Foundation, either version 3 of the License, or			*
'*		(at your option) any later version.											*
'*																					*
'*		This program is distributed in the hope that it will be useful,				*
'*		but WITHOUT ANY WARRANTY; without even the implied warranty of				*
'*		MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the				*
'*		GNU General Public License for more details.								*
'*																					*
'*		You should have received a copy of the GNU General Public License			*
'*		along with this program.  If not, see <http://www.gnu.org/licenses/>.		*
'*																					*
'*																					*
'*																					*
'************************************************************************************







'First we need to know how many days in each month
dim months(11)  'computers usually start counting at 0, so each month is going to be assigned
				'the number one less than you'd expect

		'30 days hath 
		months(8)=30	'september
		months(3)=30	'april
		months(5)=30	'june
		months(10)=30	'and november
		'all the rest have 31
		months(0)=31
		months(2)=31
		months(4)=31
		months(6)=31
		months(7)=31
		months(9)=31
		months(11)=31
		'except that asshole february
		months(1)=28
		
		'We'll deal with leap years later.

'Set your settings

'URL for the file
strBaseURL = "http://www.wunderground.com/weatherstation/WXDailyHistory.asp?ID=KPAWYNNE5"
strDayPrefix = "&day="
strMonthPrefix = "&month="
strYearPrefix = "&year="
strSuffix = "&graphspan=day&format=1"
strHDLocation = "c:\Weather Data\" 'Make sure this folder exists first

dim chosenmonth
dim chosenyear
chosenmonth = InputBox("Enter month number")
chosenyear = inputbox("Enter the year")

if 0<chosenmonth and chosenmonth<13 and chosenyear>2000 and chosenyear>=year(now())then 
	days = months(chosenmonth-1)
	else 
	msgbox("Invalid Date")
	end if


'Now let's deal with leap years
If chosenmonth=2 and (chosenyear Mod 4)=0 then
	days = days+1
end if
	
for i=1 to days
	strInputFile = strBaseURL & strDayPrefix & i & strMonthPrefix & chosenmonth & strYearPrefix & chosenyear & strSuffix
	strOutputFile = strHDLocation & "\" & chosenmonth & "_" & i & "_"& chosenyear & ".txt"

' This is all of the downloading the file stuff
' Don't mess with this part
Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

	objXMLHTTP.open "GET", strInputFile, false
	objXMLHTTP.send()

	If objXMLHTTP.Status = 200 Then
	Set objADOStream = CreateObject("ADODB.Stream")
	objADOStream.Open
	objADOStream.Type = 1 'adTypeBinary

	objADOStream.Write objXMLHTTP.ResponseBody
	objADOStream.Position = 0	'Set the stream position to the start

	Set objFSO = Createobject("Scripting.FileSystemObject")
	If objFSO.Fileexists(strOutputFile) Then objFSO.DeleteFile strOutputFile
	Set objFSO = Nothing

	objADOStream.SaveToFile strOutputFile
	objADOStream.Close
	Set objADOStream = Nothing
	End if

Set objXMLHTTP = Nothing


	'So the files aren't actually text files, they're html files
	'So instead of line breaks, they use the html code for a line
	'break, which is <br>. So we're going to read the file, and 
	'replace every instance of <br> with an actual line break.
	
	'First open the file we just saved
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(strOutputFile,1)
	'And load it's contents into a variable
	strFileText = objFileToRead.ReadAll()
	'Then replace the <br>s with vbCrLf, which is VB's preset line break character
	strFileTextFixed = Replace(strFileText, "<br>", vbCrLf)
	
	'Then overwrite the file
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(strOutputFile,True)
	objFile.Write strFileTextFixed
	objFile.Close

Next