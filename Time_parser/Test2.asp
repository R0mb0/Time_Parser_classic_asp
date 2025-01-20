<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="time_parser.class.asp"-->
<%

    Dim my_time
    Dim temp 

    set my_time = new timeParser

    temp = "16:15:13"
    Response.write("Original Time: " & temp)
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h"))
%> 
