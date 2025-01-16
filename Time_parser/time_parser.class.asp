<% 
class listOutDates

    ' Initialization and destruction'
	Sub class_initialize()
	End Sub
	
	Sub class_terminate()
	End Sub

	'Function to transform a string into array
	Private Function string_to_array(text)
		Dim length
		length = Len(text)
		Dim outArray() 
		Redim outArray(length)
		Dim index 
		For index = 0 to length - 1
    		outArray(index) = Left(Right(text,(length - index)), (1))
		Next 
		stringToArray = outArray
	End Function

	'Private Function to recognize the character that divide the time
	Private Function recognize_character(time)
		If InStr(0, time, ".") <> 0 Then 
			recognize_character = "."
			Exit Function 
		End if
		If InStr(0, time, ",") <> 0 Then 
			recognize_character = ","
			Exit Function
		End if
		If InStr(0, time, ":") <> 0 Then 
			recognize_character = ":"
			Exit Function
		End if
		If InStr(0, time, ";") <> 0 Then 
			recognize_character = ";"
			Exit Function
		End if
		If InStr(0, time, "`") <> 0 Then 
			recognize_character = "`"
			Exit Function
		End if
		If InStr(0, time, "/") <> 0 Then 
			recognize_character = "/"
			Exit Function
		End if
		If InStr(0, time, "\") <> 0 Then 
			recognize_character = "\"
			Exit Function
		End if
		If InStr(0, time, "|") <> 0 Then 
			recognize_character = "|"
			Exit Function
		End if
		If InStr(0, time, "_") <> 0 Then 
			recognize_character = "_"
			Exit Function
		End if
		If InStr(0, time, "-") <> 0 Then 
			recognize_character = "-"
			Exit Function
		End if
		If InStr(0, time, "~") <> 0 Then 
			recognize_character = "~"
			Exit Function
		End if
		If InStr(0, time, "!") <> 0 Then 
			recognize_character = "!"
			Exit Function
		End if 
		If InStr(0, time, "@") <> 0 Then 
			recognize_character = "@"
			Exit Function
		End if
		If InStr(0, time, "#") <> 0 Then 
			recognize_character = "#"
			Exit Function
		End if
		If InStr(0, time, "$") <> 0 Then 
			recognize_character = "$"
			Exit Function
		End if
		If InStr(0, time, "%") <> 0 Then 
			recognize_character = "%"
			Exit Function
		End if
		If InStr(0, time, "^") <> 0 Then 
			recognize_character = "^"
			Exit Function
		End if
		If InStr(0, time, "&") <> 0 Then 
			recognize_character = "&"
			Exit Function
		End if
		If InStr(0, time, "*") <> 0 Then 
			recognize_character = "*"
			Exit Function
		End if
		If InStr(0, time, "(") <> 0 Then 
			recognize_character = "("
			Exit Function
		End if
		If InStr(0, time, ")") <> 0 Then 
			recognize_character = ")"
			Exit Function
		End if
		If InStr(0, time, "+") <> 0 Then 
			recognize_character = "+"
			Exit Function
		End if
		If InStr(0, time, "=") <> 0 Then 
			recognize_character = "="
			Exit Function
		End if
		If InStr(0, time, "{") <> 0 Then 
			recognize_character = "{"
			Exit Function
		End if
		If InStr(0, time, "[") <> 0 Then 
			recognize_character = "["
			Exit Function
		End if
		If InStr(0, time, "}") <> 0 Then 
			recognize_character = "}"
			Exit Function
		End if
		If InStr(0, time, "]") <> 0 Then 
			recognize_character = "]"
			Exit Function
		End if
		If InStr(0, time, "'") <> 0 Then 
			recognize_character = "'"
			Exit Function
		End if
		If InStr(0, time, "<") <> 0 Then 
			recognize_character = "<"
			Exit Function
		End if
		If InStr(0, time, ">") <> 0 Then 
			recognize_character = ">"
			Exit Function
		End if
		If InStr(0, time, "e") <> 0 Then 
			recognize_character = "e"
			Exit Function
		End if
		If InStr(0, time, "and") <> 0 Then 
			recognize_character = "and"
			Exit Function
		End if
		If InStr(0, time, "()") <> 0 Then 
			recognize_character = "()"
			Exit Function
		End if
		If InStr(0, time, "[]") <> 0 Then 
			recognize_character = "[]"
			Exit Function
		End if
		If InStr(0, time, "{}") <> 0 Then 
			recognize_character = "{}"
			Exit Function
		End if
	End Function 

	'Function to divide a compact time
	Private Function split_compact_time(time, selector)
		Dim temp 
		Select case selector
			Case "h"
				Select Case Len(time)
					Case 1
						split_compact_time = "0" & time & ":00:00"
						Exit Function
					Case 2
						split_compact_time = time & ":00:00"
						Exit Function
					Case 3
						temp = string_to_array(time)
						split_compact_time = temp(0) & temp(1) & ":" & "0" & temp(2) & ":00"
						Exit Function
					Case 4
						temp = string_to_array(time)
						split_compact_time = temp(0) & temp(1) & ":" & temp(2) & temp(3) & ":00"
						Exit Function
					Case 5
						temp = string_to_array(time)
						split_compact_time = temp(0) & temp(1) & ":" & temp(2) & temp(3) & ":0" & temp(4)
						Exit Function
					Case 6
						temp = string_to_array(time)
						split_compact_time = temp(0) & temp(1) & ":" & temp(2) & temp(3) & ":" & temp(4) & temp(5)
						Exit Function
					Case Else
						Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time")
				End Select 
			Case "m"
				Select Case Len(time)
					Case 1
						split_compact_time = "00:0" & time & ":00"
						Exit Function
					Case 2
						split_compact_time = "00:" & time & ":00"
						Exit Function
					Case 3
						temp = string_to_array(time)
						split_compact_time = "00:" & temp(0) & temp(1) & ":0" & temp(2) 
						Exit Function
					Case 4
						temp = string_to_array(time)
						split_compact_time = "00:" & temp(0) & temp(1) & ":" & temp(2) & temp(3)
						Exit Function
					Case 5
						temp = string_to_array(time)
						split_compact_time = "0" & temp(0) & ":" & temp(1) & temp(2) & ":" temp(3) & temp(4)
						Exit Function
					Case 6
						temp = string_to_array(time)
						split_compact_time = temp(0) & temp(1) & ":" & temp(2) & temp(3) & ":" & temp(4) & temp(5) 
						Exit Function
					Case Else
						Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time")
				End Select 
			Case "s"
				Select Case Len(time)
					Case 1
						split_compact_time = "00:00:0" & time
						Exit Function
					Case 2
						split_compact_time = "00:00:" & time
						Exit Function
					Case 3
						temp = string_to_array(time)
						split_compact_time = "00:0" & temp(0) & ":" temp(1) & temp(2)
						Exit Function
					Case 4
						temp = string_to_array(time)
						split_compact_time = "00:" & temp(0) & temp(1) & ":" temp(2) & temp(3)
						Exit Function
					Case 5
						temp = string_to_array(time)
						split_compact_time = "0" & temp(0) & ":" & temp(1) & temp(2) & ":" temp(3) & temp(4)
						Exit Function
					Case 6
						temp = string_to_array(time)
						split_compact_time = temp(0) & temp(1) & ":" & temp(2) & temp(3) & ":" & temp(4) & temp(5) 
					Case Else
						Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time")
				End Select 
			Case Else
				Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad selector")
		End Select
	End Function 

	'Selector: 
	' h = Hours
	' m = Minutes
	' s = Seconds
    Public Function time_parser(time, selector)
		Dim character
		character = recognize_character(time)
		If Len(character) <> 0 Then 
			Select Case Len(time) - len(replace(time, character, ""))
				Case 1
					Select Case selector
						Case "h"
						Case "m"
						Case "s"
					End Select
				Case 2
					time_parser = replace(time, character, ":")
					Exit Function 
				Case Else 
					Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time")
			End Select
		Else
			time_parser = split_compact_time(time, selector)
		End If
    End Function 
End Class 
%> 