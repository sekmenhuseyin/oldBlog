<!--#include file="uploader.asp"--><%if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then response.Redirect(SiteAdres&"/")
Set objUpload=New clsUpload:pass=0
strFileName=Server.Mappath("\HuseyinSekmenoglu\upload")&"\"&year(date())&month(date())&day(date())&hour(time())&minute(time())&second(time())&".jpg"&objUpload.Fields("File1").FileName:objUpload("File1").SaveAs strFileName
Set objUpload=Nothing
response.redirect "/HuseyinSekmenoglu/DosyaYukle/"&pass
%>