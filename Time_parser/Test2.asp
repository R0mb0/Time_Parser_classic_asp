<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="time_parser.class.asp"-->
<%

    Dim my_time
    Dim temp 

    set my_time = new timeParser

    temp = "16:15:13"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "16_15_13"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "16_15"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")
%> 
