<%
Private Function EscapeSingleQuotes(byVal strInput)
	If Not IsNull(strInput) Then
		Dim strOutput
		strOutput = Replace(strInput, "'", "\'")
		EscapeSingleQuotes = strOutput
	Else
		EscapeSingleQuotes = ""
	End If
End Function

Private Function EscapeDoubleQuotes(byVal strInput)
	If Not IsNull(strInput) Then
		Dim strOutput
		strOutput = Replace(strInput, chr(34), "&quot;")
		EscapeDoubleQuotes = strOutput
	Else
		EscapeDoubleQuotes = ""
	End If
End Function

Private Function EscapeQuotes(byVal strInput)
	If Not IsNull(strInput) Then
		Dim strOutput
		strOutput = Replace(strInput, "'", "\'")
		strOutput = Replace(strOutput, chr(34), "&quot;")
		EscapeQuotes = strOutput
	Else
		EscapeQuotes = ""
	End If
End Function
%>