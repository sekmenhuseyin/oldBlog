<!--#include file="tasarim.asp"--><%top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
if ubound(SayfaAdres)=2 then
	if SayfaAdres(2)="tamam" then
		response.Write "<div class=""success"">Mesajýnýz bana ulaþtýrýldý. En kýsa zamanda size geri cevap yazmaya çalýþacaðým.</div>"
	elseif SayfaAdres(2)="hata" then
		Response.Write "<div class=""error"">Bir þeyleri eksik yazmýþsýnýz. Kusura bakmayýn.<br />Tekrar denerseniz sevinirim.</div>"
	end if
end if
%><h1>Ýletiþim</h1><a href="mailto:<%=EPosta%>" title="Bana mail at"><%=EPosta%></a>, <a href="http://feeds.feedburner.com/HuseyinSekmenoglu" title="RSS" target="_blank">RSS</a><br /><br /><br /><%
Response.Write "<div id=""iletisim-mesajyaz"" class=""Mesaj""><h2>Mesaj At</h2><br />"&_
"<form action="""&SiteAdres&"/islem.asp"" method=""post"" id=""mesajyazmaformu""><br />"&_
"<label for=""ekleyen"">Ýsim (*)&nbsp;<input type=""text"" name=""ekleyen"" class=""TextField"" value="""&Z1&""" /></label><br />"&_
"<label for=""eposta"">e-Posta (*)&nbsp;<input type=""text"" name=""eposta"" class=""TextField"" value="""&Z2&""" /></label><br />"&_
"<label for=""metin"">Mesajýnýz (*)&nbsp;</label><textarea name=""metin"" rows=""8"" cols=""30""></textarea><br /><br />"&_
"Güvenlik (*)&nbsp;<br />Aþaðýdaki yazýnýn aynýsýný yan tarafa yaz ki insan olduðunu anlayayým<label for=""guvenlik"">"&_
"<img src="""&SiteAdres&"/Contents/captcha.asp"" alt=""This Is CAPTCHA Image"" id=""captcha"" />&nbsp;<a href=""javascript:RefreshImage('captcha')"" style=""font-size:smaller"">Change image</a>"&_
"<input type=""text"" name=""guvenlik"" class=""TextField"" value="""" /></label><br />"&_
"<input name=""BanaMesajGonder"" type=""submit"" value=""Gönder"" /></form></div>"
Response.Write "<div id=""iletisim-messsenger"" class=""Mesaj""><h2>Messenger ile Konuþ</h2><br />"&_
"<iframe src=""http://settings.messenger.live.com/Conversation/IMMe.aspx?invitee=1d076185e6d38029@apps.messenger.live.com&mkt=tr-TR"" width=""400"" height=""350"""&_
" style=""border:none;width:400px;height:350px;"" frameborder=""0"" scrolling=""no""></iframe></div>"
bottom%>