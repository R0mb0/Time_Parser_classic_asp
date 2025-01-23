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

    temp = "16:15:13"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "16_15_13"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "16:15:13"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "16_15_13"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "16:15:3"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "16_15_3"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "16:15:3"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "16_15_3"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "16:15:3"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "16_15_3"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "161513"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "161513"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "161513"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "16_15"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "16:15"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "16-5"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")
    
    temp = "16_15"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "16:15"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "16-5"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "16_15"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "16:15"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "16-5"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "1615"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "165"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")
    
    temp = "1615"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "165"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "1615"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "165"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "16"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "1:6"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "h") & " Setting: h <br><br>")

    temp = "16"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "1-6"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "m") & " Setting: m <br><br>")

    temp = "16"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")

    temp = "1_6"
    Response.write("Original Time: " & temp & "<br>")
    Response.write("Time Parsed: " & my_time.time_parser(temp, "s") & " Setting: s <br><br>")
%> 
