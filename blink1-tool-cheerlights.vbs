Set colDict = CreateObject("Scripting.Dictionary")
colDict.Add "Red", "FF0000"
colDict.Add "Orange", "FF9C00"
colDict.Add "Yellow", "FFDE00"
colDict.Add "Green", "00FF00"
colDict.Add "Cyan", "00FFFF"
colDict.Add "Blue", "0000FF"
colDict.Add "Magenta", "FF00DE"
colDict.Add "Purple", "CD00FF"
colDict.Add "Warmwhite", "FFEBAF"
colDict.Add "White", "FFF4F0"
colDict.Add "Pink", "FFDEDE"
colDict.Add "Oldlace", "FFEBAF"
colKeys = colDict.Keys

colCheer = getCheerlights

For Each strKey in colKeys
	If LCase(colCheer) = LCase(strKey) Then
		echo "Setting cheerlights to " & strKey & " " & colDict.Item(strKey)
		Call blink1(colDict.Item(strKey))
		Exit For
	End If
Next

'################################
Function blink1(color)
	r = CLng("&h" & Left(color, 2))
	g = CLng("&h" & Mid(color, 3, 2))
	b = CLng("&h" & Right(color, 2))
	echo r & " " & g & " " & b
	Set WshShell = WScript.CreateObject("WScript.Shell")
	return = WshShell.Run("blink1-tool.exe -m 100 --rgb " & r & "," & g & "," & b, 0, true)
	Set WshShell = Nothing
End Function

Function getCheerlights
	strUrl = "http://api.thingspeak.com/channels/1417/field/1/last.txt"
	Set http = CreateObject("Microsoft.XmlHttp")
	http.Open "GET", strUrl, False
	http.Send
	If http.Status = 200 Then
		getCheerlights = http.responseText
	End If
	Set http = Nothing
End Function

Function echo(strText)
	If False Then WScript.Echo strText
End Function

'################################
