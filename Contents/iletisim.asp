<!--#include file="tasarim.asp"--><%top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
if ubound(SayfaAdres)=2 then
	if SayfaAdres(2)="tamam" then
		response.Write "<div class=""success"">Mesaj�n�z bana ula�t�r�ld�. En k�sa zamanda size geri cevap yazmaya �al��aca��m.</div>"
	elseif SayfaAdres(2)="hata" then
		Response.Write "<div class=""error"">Bir �eyleri eksik yazm��s�n�z. Kusura bakmay�n.<br />Tekrar denerseniz sevinirim.</div>"
	end if
end if
%><h1>�leti�im</h1><a href="mailto:<%=EPosta%>" title="Bana mail at"><%=EPosta%></a>, <a href="http://feeds.feedburner.com/HuseyinSekmenoglu" title="RSS" target="_blank">RSS</a><br /><br /><br /><%
Response.Write "<div id=""iletisim-mesajyaz"" class=""Mesaj""><h2>Mesaj At</h2><br />"&_
"<form action="""&SiteAdres&"/islem.asp"" method=""post"" id=""mesajyazmaformu""><br />"&_
"<label for=""ekleyen"">�sim (*)&nbsp;<input type=""text"" name=""ekleyen"" class=""TextField"" value="""&Z1&""" /></label><br />"&_
"<label for=""eposta"">e-Posta (*)&nbsp;<input type=""text"" name=""eposta"" class=""TextField"" value="""&Z2&""" /></label><br />"&_
"<label for=""metin"">Mesaj�n�z (*)&nbsp;</label><textarea name=""metin"" rows=""8"" cols=""30""></textarea><br /><br />"&_
"G�venlik (*)&nbsp;<br />A�a��daki yaz�n�n ayn�s�n� yan tarafa yaz ki insan oldu�unu anlayay�m<label for=""guvenlik"">"&_
"<img src="""&SiteAdres&"/Contents/captcha.asp"" alt=""This Is CAPTCHA Image"" id=""captcha"" />&nbsp;<a href=""javascript:RefreshImage('captcha')"" style=""font-size:smaller"">Change image</a>"&_
"<input type=""text"" name=""guvenlik"" class=""TextField"" value="""" /></label><br />"&_
"<input name=""BanaMesajGonder"" type=""submit"" value=""G�nder"" /></form></div>"
Response.Write "<div id=""iletisim-messsenger"" class=""Mesaj""><h2>Messenger ile Konu�</h2><br />"&_
"<iframe src=""http://settings.messenger.live.com/Conversation/IMMe.aspx?invitee=1d076185e6d38029@apps.messenger.live.com&mkt=tr-TR"" width=""400"" height=""350"""&_
" style=""border:none;width:400px;height:350px;"" frameborder=""0"" scrolling=""no""></iframe></div>"
bottom%>