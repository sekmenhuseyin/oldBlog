<%ua=Request.ServerVariables("HTTP_USER_AGENT"):if len(trim(ua))<10 then response.end'just for security
If Instr(ua, "MSIE") Then
	Browser="Microsoft Internet Explorer":mobile_browser=false'not a mobile
	If Instr(ua, "MSIE 9.") Then
		Browser="Microsoft Internet Explorer 9"
	ElseIf Instr(ua, "MSIE 8.") Then
		Browser="Microsoft Internet Explorer 8"
	ElseIf Instr(ua, "MSIE 7.") Then
		Browser="Microsoft Internet Explorer 7"
	ElseIf Instr(ua, "MSIE 6.") Then
		Browser="Microsoft Internet Explorer 6"
	ElseIf Instr(ua, "MSIE 5.") Then
		Browser="Microsoft Internet Explorer 5"
	Elseif Instr(ua, "MSIE 4.") Then
		Browser="Microsoft Internet Explorer 4"
	Elseif Instr(ua, "MSIE 3.") Then
		Browser="Microsoft Internet Explorer 3"
	End If
ElseIf Instr(ua, "Firefox") Then
	Browser="Mozilla Firefox":mobile_browser=false'not a mobile
	If Instr(ua, "Firefox/4") Then
		Browser="Mozilla Firefox 4"
	Elseif Instr(ua, "Firefox/3") Then
		Browser="Mozilla Firefox 3"
	Elseif Instr(ua, "Firefox/2") Then
		Browser="Mozilla Firefox 2"
	End If
ElseIf Instr(ua, "Chrome") Then
	Browser="Google Chrome":mobile_browser=false'not a mobile
	If Instr(ua, "Chrome/4") Then
		Browser="Google Chrome 4"
	ElseIf Instr(ua, "Chrome/3") Then
		Browser="Google Chrome 3"
	ElseIf Instr(ua, "Chrome/2") Then
		Browser="Google Chrome 2"
	End If
ElseIf Instr(ua, "Safari") Then
	Browser="Safari":mobile_browser=false'not a mobile
ElseIf Instr(ua, "Opera") Then
	Browser="Opera":mobile_browser=false'not a mobile
	If Instr(ua, "Opera Mini") Then Browser="Opera Mini":mobile_browser=true'mobile
End If
If Browser="" OR IsNull(Browser) OR Len(Browser) < 3 Then Browser="Unknown - "&ua:mobile_browser=false'not a mobile

If Instr(ua, "Windows") or Instr(ua, "Win") Then
	os="Windows":mobile_browser=false'not a mobile
	if Instr(ua, "Windows NT 6.1;") Then
		os="Windows Seven"
	Elseif Instr(ua, "Windows NT 6.0;") Then
		os="Windows Vista"
	Elseif Instr(ua, "Windows NT 5") Then
		os="Windows XP"
	Elseif Instr(ua, "Windows NT 4.") or Instr(ua, "Windows 9") Then
		os="Windows 9x"
	Elseif Instr(ua, "Windows CE") Then
		os="Windows CE":mobile_browser=true'mobile
	End if
Elseif Instr(ua, "Mac") Then
	os="Mac":mobile_browser=false'not a mobile
Elseif Instr(ua, "Linux") Then
	os="Linux":mobile_browser=false'not a mobile
Elseif Instr(ua, "Android") Then
	os="Android":mobile_browser=true'mobile
Elseif Instr(ua, "X11") Then
	os="UNIX":mobile_browser=false'not a mobile
Elseif Instr(ua, "Iphone") Then
	os="Iphone":mobile_browser=true'mobile
Elseif Instr(ua, "BlackBerry") Then
	os="BlackBerry":mobile_browser=true'mobile
Elseif Instr(ua, "Nokia") Then
	os="Nokia":mobile_browser=true'mobile
Elseif Instr(ua, "LG") Then
	os="LG":mobile_browser=true'mobile
Elseif Instr(ua, "HTC") Then
	os="HTC":mobile_browser=true'mobile
Elseif Instr(ua, "Motorola") Then
	os="Motorola":mobile_browser=true'mobile
Elseif Instr(ua, "Samsung") Then
	os="Samsung":mobile_browser=true'mobile
Elseif Instr(ua, "Ericsson") or Instr(ua, "Sony") Then
	os="SonyEricsson":mobile_browser=true'mobile
End If
if Instr(ua, "PlayStation") then os="PlayStation":Browser="PlayStation":mobile_browser=false
If os="" OR IsNull(os) OR Len(os) < 3 Then os="Unknown - "&ua:mobile_browser=false'not a mobile%>