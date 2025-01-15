<% 
class listOutDates

    ' Initialization and destruction'
	Sub class_initialize()
	End Sub
	
	Sub class_terminate()
	End Sub

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

    Public Function time_parser(time, selector)

    End Function 
End Class 
%> 