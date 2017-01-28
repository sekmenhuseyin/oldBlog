<% 
Function Temizle(gelenveri)
  gelenveri = Replace(gelenveri ,"<","&lt;",1,-1,1)
  gelenveri = Replace(gelenveri ,">","&gt;",1,-1,1)
  gelenveri = Replace(gelenveri ,Chr(34),"&#34;",1,-1,1)
  gelenveri = Replace(gelenveri ,Chr(39),"&#39;",1,-1,1)
  gelenveri = Replace(gelenveri ,Chr(13),"<br />",1,-1,1)
  gelenveri = Replace(gelenveri ,"[code]","<div id=kod>")
  gelenveri = Replace(gelenveri ,"[/code]","</div>")
Temizle = gelenveri
End Function 

Public Function Link(byVal Text) 
     Set objReg = New RegEXP 
          objReg.Global = True 
          objReg.IgnoreCase = True 
          objReg.Pattern = "\[link:\s*(.+?)\]\s*(.+?)\[/link]" 
     Text = objReg.Replace(Text,"<a href=""$1"" target=""_blank"">$2</a>") 
Link = Text 
End Function

Public Function Resim(byVal Text) 
     Set objReg = New RegEXP 
          objReg.Global = True 
          objReg.IgnoreCase = True 
          objReg.Pattern = "\[resim]\s*(.+?)\[/resim]" 
     Text = objReg.Replace(Text,"<div id=""resim""><img src=""$1"" border=""0"" align=""left"" style=""border: 3px double #C0C0C0;""></div>") 
Resim = Text 
End Function

Public Function Google(byVal Text) 
     Set objReg = New RegEXP 
          objReg.Global = True 
          objReg.IgnoreCase = True 
          objReg.Pattern = "\[ara]\s*(.+?)\[/ara]" 
     Text = objReg.Replace(Text,"<a href=""http://www.google.com.tr/search?hl=tr&q=$1"" target=""_blank"">$1</a>") 
Google = Text 
End Function

Public Function Video(byVal Text) 
     Set objReg = New RegEXP 
          objReg.Global = True 
          objReg.IgnoreCase = True 
          objReg.Pattern = "\[video]\s*(.+?)\[/video]" 
     Text = objReg.Replace(Text,"<object width=""400"" height=""334""><param name=""movie"" value=""http://www.youtube.com/v/$1""></param><param name=""wmode"" value=""transparent""></param><embed src=""http://www.youtube.com/v/$1"" type=""application/x-shockwave-flash"" wmode=""transparent"" width=""400"" height=""334""></embed></object>") 
Video = Text 
End Function
%><!-- #Include File = "abt-yapilandirma.asp" --><%
response.ContentType="text/xml"
response.write "<?xml version=""1.0"" encoding=""windows-1254""?>"
response.write "<rss version=""2.0""><channel>"

response.write "<title>"&SiteAdi&" "&SiteBaslik&"</title>"
response.write "<link>"&SiteAdres&"</link>"
response.write "<description>ABT Blog v1.0</description>"
response.write "<webMaster>"&YoneticiEposta&"</webMaster>"
response.write "<copyright>"&YoneticiAdi&"</copyright>"
response.write "<language>tr-TR</language>"

Set ObjRs = Server.CreateObject ("ADODB.RecordSet")
ObjRs.open "select * from blog order by id desc;", ObjCon, 1, 3

for i = 1 to 10
if ObjRs.eof then exit for
instr(tmpObje("metin"),"<hr />")>0 then Bolunmus=left(tmpObje("metin"),instr(tmpObje("metin"),"<hr />")-1) else Bolunmus=tmpObje("metin")
response.write "<item>"
response.write "<title>"&ObjRs("baslik")&"</title>"
response.write "<pubDate>" & GetRFC822Date(ObjRs("tarih")) & "</pubDate>"
response.write "<link>"&SiteAdres&"/blog/"&ObjRs("id")&"/"&SEO_Olustur(ObjRs("baslik"))&".html</link>"
response.write "<description><![CDATA["&Video(Google(Resim(Link(Temizle(Bolunmus)))))&"]]></description>"
response.write "</item>"
ObjRs.Movenext
Next

response.write "</channel></rss>"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetRFC822Date
' Purpose:
' Gets date in the RFC 822 format (with 4-digit year) as required by the RSS
' 2.0 spec.
' Parameters:
' dtmDate - Date to format.
' Returns:
' Formatted date.
' Revisions:
' [tempus_rook@hotmail.com 2003-11-02] Code written.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetRFC822Date(dtmDate)
    ' XXX what about daylight savings?
    Dim strDate
    strDate = Left(WeekdayName(DatePart("w", dtmDate)), 3) & ", " & _
        LeadingZero(DatePart("d", dtmDate))
     strDate = strDate & _
        " " & Left(MonthName(DatePart("m", dtmDate)), 3) & " " & _
        DatePart("yyyy", dtmDate) & _
        " " & LeadingZero(DatePart("h", dtmDate))
     strDate = strDate & ":" & LeadingZero(DatePart("n", dtmDate)) & _
        ":" & LeadingZero(DatePart("s", dtmDate)) & " GMT"
    GetRFC822Date = strDate
    ' Format: Sun, 02 Nov 2003 13:40:01 GMT
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LeadingZero
' Purpose:
' Adds a leading zero to single digit numbers.
' Parameters:
' strText - Var to add a zero to, if necessary.
' Returns:
' Original text, with leading zero if required.
' Revisions:
' [tempus_rook@hotmail.com 2003-11-02] Code written.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function LeadingZero(strText)
    If Len(strText) = 1 Then
        LeadingZero = "0" & strText
    Else
        LeadingZero = strText
    End If
End Function
%>