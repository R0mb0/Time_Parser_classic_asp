<% 
class listOutDates

    ' Initialization and destruction'
	Sub class_initialize()
	End Sub
	
	Sub class_terminate()
	End Sub

	'Check a string is a time
	Private Function IsTime(str)
  		If str = "" Then
    		IsTime = false
  		Else
    		On Error Resume Next
    		TimeValue(str)
    		If Err.number = 0 Then
      			IsTime = true
    		Else
      			IsTime = false
    		End if
    		On Error GoTo 0
  		End if
	end function

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

	'Function to check if there are two identical characters 
	Private Function check_consistency(time)
		Dim character
		character = recognize_character(time)
		check_character_position(time,character) 'Checking 
		If Len(character) <> 0 Then 
			If Len(time) >= 5 Then 
				Dim temp 
				temp = Right(time, Len(time) - InStr(time, character))
				check_character_position(temp, character) 'Checking 
				Dim second_character
				second_character = recognize_character(temp)
				If character = second_character Then 
					check_consistency = character
				Else
					check_consistency = replace(time, character, "")
				End If 
			Else
				check_consistency = character
				Exit Function
			End If 
		Else
			check_consistency = character
			Exit Function
		End If 
	End Function 

	'Function to check the position of the character 
	Private Function check_character_position(time, character)
		Dim position 
		position = InStr(time, character)
		Dim arr_time
		arr_time = string_to_array(time)
		If arr_time(0) = character and 0 = position - 1 Then 
			Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - character position")
		End If 
		If arr_time(UBound(arr_time)) = character and UBound(arr_time) = position - 1 Then 
			Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - character position")
		End If
		If arr_time(position - 1) = character and arr_time(position) = character Then 
			Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - character position")
		End If 
	End Function 

	'Function to divide a compact time
	Private Function split_compact_time(time, selector)
		Dim temp 
		Select Case LCase(selector)
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
						Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - Impossible to Split")
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
						Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - Impossible to Split")
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
						Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - Impossible to Split")
				End Select 
			Case Else
				Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad selector - Impossible to Split")
		End Select
	End Function 

	'Selector: 
	' h = Hours
	' m = Minutes
	' s = Seconds
	' TIME MUST BE IN 24H FORMAT
    Public Function time_parser(time, selector)
		Dim character
		Dim time_contracted
		Dim my_time
		Dim temp 
		character = check_consistency(time) 
		'Check if is present a special character
		If Len(character) <> 0 Then
			my_time = time
			'Check if character is a fixed time from check_consistency
			If Len(time) = Len(character) Then 
				my_time = character
				character = recognize_character(my_time)
			End If
			time_contracted = replace(my_time, character, "")
			'Check the number of the characters 
			Select Case Len(my_time) - Len(my_time)
				Case 1
					'Check the number of digits of the time without the character
					Select case Len(time_contracted)
						Case 2
							'Check if the character is between the digits
							If InStr(my_time,character) = 2 Then 
								'Check the selector value
								Select Case LCase(selector)
									Case "h"
										temp = string_to_array(my_time)
										time_parser = "0" temp(0) & ":0" & temp(1) & ":00" 
										Exit Function
									Case "m"
										temp = string_to_array(my_time)
										time_parser = "00:0" temp(0) & ":0" & temp(1) 
										Exit Function
									Case "s"
										temp = string_to_array(my_time)
										time_parser = "00:0" temp(0) & ":0" & temp(1)
										Exit Function
									Case Else
										Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad selector - For parser") 
								End Select
							Else
								Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - For parser")
							End If 
						Case 3
							'Check the position of the character
							Select Case InStr(my_time,character)
								Case 2
									'Check the selector value
									Select Case LCase(selector)
          								Case "h"
											temp = string_to_array(my_time)
											time_parser = "0" & temp(0) & ":" & temp(1) & temp(2) & ":00"
          									Exit Function
          								Case "m"
											temp = string_to_array(my_time)
											time_parser = "00:0" & temp(0) & ":" & temp(1) & temp(2)
          									Exit Function
          								Case "s"
											temp = string_to_array(my_time)
											time_parser = "00:0" & temp(0) & ":" & temp(1) & temp(2)
          									Exit Function
          								Case Else
          									Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad selector - For parser") 
          							End Select
								Case 3
									'Check the selector value
									Select Case LCase(selector)
          								Case "h"
											temp = string_to_array(my_time)
											time_parser = temp(0) & temp(1) & ":0" & temp(2) & ":00"
          									Exit Function
          								Case "m"
											temp = string_to_array(my_time)
											time_parser = "00:" & temp(0) & temp(1) & ":0" & temp(2)
          									Exit Function
          								Case "s"
											temp = string_to_array(my_time)
											time_parser = "00:" & temp(0) & temp(1) & ":0" & temp(2)
          									Exit Function
          								Case Else
          									Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad selector - For parser") 
          							End Select
								Case Else 
									Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - For parser") 
							End Select
						Case 4
							Select Case LCase(selector)
								Case "h"
									temp = string_to_array(my_time)
									time_parser = temp(0) & temp(1) & ":" & temp(2) & temp(3) & ":00"
									Exit Function
								Case "m"
									temp = string_to_array(my_time)
									time_parser = "00:" & temp(0) & temp(1) & ":" & temp(2) & temp(3) 
									Exit Function
								Case "s"
									temp = string_to_array(my_time)
									time_parser = "00:" & temp(0) & temp(1) & ":" & temp(2) & temp(3)
									Exit Function
								Case Else
									Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad selector - For parser") 
							End Select
						Case Else 
							Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - For parser")
					End Select
				'Case ther're two character
				Case 2
					'Check the number of digits of the time without the character
					Select Case Len(time_contracted)
						Case 5
							temp = string_to_array(my_time)
							time_parser = "0" & temp(0) & ":" & temp(1) & temp(2) & ":" temp(3) & temp(4)
							Exit Function 
						Case 6 
							temp = string_to_array(my_time)
							'------------------------------------------------------------- Here to complete 
							time_parser = temp(0) & temp(1) & ":" & temp(2) & temp(3) & ":" temp(4) & temp(5)
							Exit Function 
						Case Else 
							Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - For parser")
					End Select 
				Case Else 
					Call Err.Raise(vbObjectError + 10, "time_parser.class", "Bad time - For parser")
			End Select
		Else
			If IsTime(split_compact_time(my_time, selector))
				time_parser = split_compact_time(my_time, selector)
			Else 
				Call Err.Raise(vbObjectError + 10, "time_parser.class", "Time is not possible")
			End If 
		End If
    End Function 
End Class 
%> 