Set colDict = CreateObject("Scripting.Dictionary")
colDict.Add "Red", "FF0000"
colDict.Add "Orange", "FF6100"
colDict.Add "Yellow", "FFB600"
colDict.Add "Green", "00FF00"
colDict.Add "Cyan", "00FFA5"
colDict.Add "Blue", "0000FF"
colDict.Add "Magenta", "FF00A6"
colDict.Add "Purple", "B600FF"
colDict.Add "Warmwhite", "FFC588"
colDict.Add "White", "FFCCBB"
colDict.Add "Pink", "FF67C2"
colDict.Add "Oldlace", "FFC588"
colKeys = colDict.Keys

For Each strKey in colKeys
	echo "Setting cheerlights to " & strKey & " " & colDict.Item(strKey)
	Call blink1RGB(colDict.Item(strKey))
	WScript.Sleep 1000
Next

'################################

Function getCheerlights
	'blink1RGB("000000")
	strUrl = "http://api.thingspeak.com/channels/1417/field/1/last.txt"
	Set http = CreateObject("Microsoft.XmlHttp")
	http.Open "GET", strUrl, False
	http.Send
	If http.Status = 200 Then
		getCheerlights = http.responseText
	End If
	Set http = Nothing
End Function

Function blink1Cheer(colCheer)
	For Each strKey in colKeys
		If LCase(colCheer) = LCase(strKey) Then
			echo "Setting cheerlights to " & strKey & " " & colDict.Item(strKey)
			Call blink1RGB(colDict.Item(strKey))
			Exit For
		End If
	Next
End Function

Function blink1RGB(colRGB)
	r = CLng("&h" & Left(colRGB, 2))
	g = CLng("&h" & Mid(colRGB, 3, 2))
	b = CLng("&h" & Right(colRGB, 2))
	echo "blink1-tool.exe -g -m 100 --rgb " & r & "," & g & "," & b
	Set WshShell = WScript.CreateObject("WScript.Shell")
	return = WshShell.Run("blink1-tool.exe -g -m 100 --rgb " & r & "," & g & "," & b, 0, true)
	Set WshShell = Nothing
End Function

Function echo(strText)
	If True Then WScript.Echo strText
End Function

'################################
