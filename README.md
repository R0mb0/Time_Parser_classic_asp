# Time Parser in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/45b733bc57964de78b9d4f046b03d12e)](https://app.codacy.com/gh/R0mb0/Time_Parser_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Time_Parser_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Time_Parser_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## `time_parser.class.asp`'s avaible functions

- Funtion to parse a time by a selector used to understand the time -> `Public Function time_parser(time, selector)`
  >
  > - <ins>Where the selector could be:</ins>
  >   - "h" -> for interpret time from Hours
  >   - "m" -> for interpret time from Minutes
  >   - "s" -> for interpret time from Seconds

## How to use 

> From `Test.asp`

1. Initialize the class
   ```
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="time_parser.class.asp"-->
   <%
      Dim my_time 
      set my_time = new timeParser
   ```
2. Parse times
   ```
     Dim temp
   
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
   %>
   ```
