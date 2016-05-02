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


Function checkInput
	'Check for any errors, assign an errorcode if necessary.
	dim errcode
	errorcode = -1
	
	'User didn't enter a year
	if(userInput.year.value) = "" then
		'No year entered
		errorcode = 1
	else 
		if(dateSerial(userInput.year.value,userinput.startMonth.value,1) >= Now) then
			'Start month is in the future
			errorcode=3
		end if
		if(dateSerial(userInput.year.value,userinput.endMonth.value,1) > Now) then
			'End month is in the future
			errorcode=4
		end if
		if(dateSerial(userInput.year.value,userinput.endMonth.value,1) = dateSerial(Year(Now), Month(Now), 1)) then
			'End month is the current month
			errorcode=5
		end if
	end if
	if(userinput.startMonth.value > userinput.endMonth.value) then
	'Start month is after end month
		errorcode = 6
	end if 
	'User didn't select an output folder
	if(userInput.saveLocation.value) = "" then
		errorcode = 2
	end if	
	
	if(radioInput(userInput.weatherStation)="") then 
		errorcode = 7
	end if
	
	if (checkErrors(errorcode)=-1) then
		call downloadFiles(userinput.startMonth.value, userinput.endMonth.value, userinput.year.value, userinput.saveLocation.value, radioInput(userinput.weatherStation))
	end if
End Function

function downloadDayFile(strStationID, intDay, intMonth, intYear, strHDLocation)
	
	'Set your settings
	'URL for the file
	strBaseURL = "http://www.wunderground.com/weatherstation/WXDailyHistory.asp?ID=KPAWYNNE5"
	strDayPrefix = "&day="
	strMonthPrefix = "&month="
	strYearPrefix = "&year="
	strSuffix = "&graphspan=day&format=1"

	
	strInputFile = strBaseURL & strDayPrefix & intDay & strMonthPrefix & intMonth & strYearPrefix & intYear & strSuffix
	strOutputFile = strHDLocation & "\" & intMonth & "_" & intDay & "_"& intYear & ".txt"

	' This is all of the downloading the file stuff
	' Don't mess with this part
	Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")

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


end function

function downloadFiles(sMonth, eMonth, eYear, saveLoc, wsID)
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


for m = sMonth to eMonth
	days = months(m-1)
	If chosenmonth=2 and (chosenyear Mod 4)=0 then
		days = 29
	end if
	for i=1 to days
		call downloadDayFile(wsID, i, m, eYear, saveLoc)
	Next
Next

end function

Function checkErrors(code)
	select case code
		Case 1
			msgbox("You must enter a valid year!")
			checkErrors = 0
		Case 2
			msgbox("You must select a folder to save the files!")
			checkErrors = 0
		Case 3
			msgbox("You can't pick a date in the future! I'm not a fortune teller!")
			checkErrors = 0
		Case 4
			msgbox("You can't pick a date in the future! I'm not a fortune teller!")
			checkErrors = 0
		Case 5
			msgbox("We're still in that month! Pick an earlier Ending Month.")
			checkErrors = 0
		Case 6
			msgBox("End month must be before start month!")
			checkErrors = 0
		Case 7
			msgBox("Pick a weather station, bruh.")
			checkErrors = 0
		Case else
			checkErrors = -1
	end select
end function


Function radioInput(radioControl)
	dim i
	for i=0 to radioControl.length - 1
		if radioControl(i).checked then
			radioInput = radioControl(i).value
		end if
	next
End Function