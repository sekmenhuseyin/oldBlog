<% 
Function Temizle(gelenveri)
  if instr(gelenveri,"<hr />")>0 then gelenveri=left(gelenveri,instr(gelenveri,"<hr />")-1)
  gelenveri=Replace(gelenveri ,"<p>","")
  gelenveri=Replace(gelenveri ,"</p>","")
  gelenveri=Replace(gelenveri ,"<","&lt;",1,-1,1)
  gelenveri=Replace(gelenveri ,">","&gt;",1,-1,1)
  gelenveri=Replace(gelenveri ,Chr(34),"&#34;",1,-1,1)
  gelenveri=Replace(gelenveri ,Chr(39),"&#39;",1,-1,1)
  gelenveri=Replace(gelenveri ,Chr(13),"<br />",1,-1,1)
  gelenveri=Replace(gelenveri ,"Ý","I")
  gelenveri=Replace(gelenveri ,"ý","i")
  gelenveri=Replace(gelenveri ,"Ð","G")
  gelenveri=Replace(gelenveri ,"ð","g")
  gelenveri=Replace(gelenveri ,"Ü","U")
  gelenveri=Replace(gelenveri ,"ü","u")
  gelenveri=Replace(gelenveri ,"Þ","S")
  gelenveri=Replace(gelenveri ,"þ","s")
  gelenveri=Replace(gelenveri ,"Ç","C")
  gelenveri=Replace(gelenveri ,"ç","c")
  gelenveri=Replace(gelenveri ,"Ö","O")
  gelenveri=Replace(gelenveri ,"ö","o")
Temizle=gelenveri
End Function 

Public Function Link(byVal Text) 
     Set objReg=New RegEXP 
          objReg.Global=True 
          objReg.IgnoreCase=True 
          objReg.Pattern="\[link:\s*(.+?)\]\s*(.+?)\[/link]" 
     Text=objReg.Replace(Text,"<a href=""$1"" target=""_blank"">$2</a>") 
Link=Text 
End Function

Public Function Resim(byVal Text) 
	Text=Replace(Text,"../upload",SiteAdres&"/upload",1,-1,1)
	Resim=Text 
End Function

Public Function Google(byVal Text) 
     Set objReg=New RegEXP 
          objReg.Global=True 
          objReg.IgnoreCase=True 
          objReg.Pattern="\[ara]\s*(.+?)\[/ara]" 
     Text=objReg.Replace(Text,"<a href=""http://www.google.com.tr/search?hl=tr&q=$1"" target=""_blank"">$1</a>") 
Google=Text 
End Function

Public Function Video(byVal Text) 
     Set objReg=New RegEXP 
          objReg.Global=True 
          objReg.IgnoreCase=True 
          objReg.Pattern="\[video]\s*(.+?)\[/video]" 
     Text=objReg.Replace(Text,"<object width=""400"" height=""334""><param name=""movie"" value=""http://www.youtube.com/v/$1""></param><param name=""wmode"" value=""transparent""></param><embed src=""http://www.youtube.com/v/$1"" type=""application/x-shockwave-flash"" wmode=""transparent"" width=""400"" height=""334""></embed></object>") 
Video=Text 
End Function
Function SEO_Olustur(BlogBaslik)
BlogBaslik=lcase(BlogBaslik)
BlogBaslik=Replace(BlogBaslik ,"!","",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"?","",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"'","",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"_","-",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"+","-",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"ý","i",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"þ","s",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"ð","g",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"ü","u",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"ö","o",1,-1,1)
BlogBaslik=Replace(BlogBaslik ,"ç","c",1,-1,1)
BlogBaslik=Replace(BlogBaslik ," ","-",1,-1,1)
if right(BlogBaslik,1)="-" then BlogBaslik=left(BlogBaslik,len(BlogBaslik)-1)
SEO_Olustur=BlogBaslik
End Function
%><!-- #Include File="Contents/abt-yapilandirma.asp" --><%
response.ContentType="application/rss+xml":Response.charset="windows-1254"
response.write "<?xml version=""1.0"" encoding=""windows-1254""?>"
response.write "<rss version=""2.0"" xmlns:atom=""http://www.w3.org/2005/Atom""><channel>"
response.write "<atom:link href="""&SiteAdres&"/rss.asp"" rel=""self"" type=""application/rss+xml"" />"

response.write "<title>"&Temizle(SiteAdi)&" "&Temizle(SiteBaslik)&"</title>"
response.write "<link>"&SiteAdres&"</link>"
response.write "<description>"&Temizle(SiteMetaDes)&"</description>"
response.write "<webMaster>"&EPosta&" ("&Temizle(YoneticiAdi)&")</webMaster>"
response.write "<copyright>"&Temizle(YoneticiAdi)&"</copyright>"
response.write "<language>tr-TR</language>"
	response.write "<image>"
	response.write "<url>http://www.gulistankosova.eu/HuseyinSekmenoglu/upload/20106121252.jpg</url>"
	response.write "<title>"&Temizle(SiteAdi)&" "&Temizle(SiteBaslik)&"</title>"
	response.write "<link>"&SiteAdres&"</link>"
	response.write "</image>"&vbcrlf

Set ObjRs=Server.CreateObject ("ADODB.RecordSet")
ObjRs.open "select * from blog where gorunur=true order by id desc;", ObjCon, 1, 3

for i=1 to 10
if ObjRs.eof then exit for
response.write "<item>"
response.write "<title>"&ObjRs("baslik")&"</title>"
response.write "<pubDate>" & GetRFC822Date(ObjRs("tarih"),ObjRs("saat")) & "</pubDate>"
response.write "<link>"&SiteAdres&"/blog/"&ObjRs("id")&"/"&SEO_Olustur(ObjRs("baslik"))&".html</link>"
response.write "<guid>"&SiteAdres&"/blog/"&ObjRs("id")&"/"&SEO_Olustur(ObjRs("baslik"))&".html</guid>"
response.write "<description><![CDATA["&Video(Google(Resim(Link(Temizle(ObjRs("metin"))))))&"]]></description>"
response.write "</item>"&vbcrlf
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
Function GetRFC822Date(dtmDate,dtmTime)
    ' XXX what about daylight savings?
    Dim strDate
    strDate=Left(WeekdayName(DatePart("w", dtmDate)), 3) & ", " & _
        LeadingZero(DatePart("d", dtmDate))
     strDate=strDate & _
        " " & Left(MonthName(DatePart("m", dtmDate)), 3) & " " & _
        DatePart("yyyy", dtmDate) & _
        " " & LeadingZero(DatePart("h", dtmTime))
     strDate=strDate & ":" & LeadingZero(DatePart("n", dtmTime)) & _
        ":" & LeadingZero(DatePart("s", dtmTime)) & " GMT"
    GetRFC822Date=strDate
    ' Format: Sun, 02 Nov 2003 13:40:01 GMT
    
End Function
Function LeadingZero(strText)
    If Len(strText)=1 Then LeadingZero="0" & strText Else LeadingZero=strText
End Function
%>