<!--#include file="tasarim.asp"--><%if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then response.Redirect(SiteAdres&"/")
call top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
if ubound(SayfaAdres)=2 then
	if SayfaAdres(2)="Yeni" or  SayfaAdres(2)="Kaydet" then
		response.Write "<div class=""success"">Linkiniz baþarýyla kaydedildi.</div>"
	else
		Response.Write "<div class=""error"">Linkiniz için baþlýk ve adres lazým</div>"
	end if
end if:Set ObjRs=Server.CreateObject("ADODB.RecordSet")
if ubound(SayfaAdres)=3 then
Guvenlik(SayfaAdres(3))
if SayfaAdres(2)="Duzenle" then
	StrSql="select * from linkler where id="&SayfaAdres(3):ObjRs.Open StrSql,ObjCon,1,3
	Response.Write "<div id=""Link-Guncelle"" class=""Mesaj"">"
	Response.Write "<form action="""&SiteAdres&"/islem.asp"" method=""post"" id=""linkform""><h2>Link Deðiþtir</h2><br /><br />"&_
	"<div><div>Baþlýk (*) </div>&nbsp;<input type=""text"" name=""linkBaslik"" class=""TextField"" value="""&ObjRs("baslik")&""" /></div>"&_
	"<div><div>Adres (*) </div>&nbsp;<input type=""text"" name=""linkAdres"" class=""TextField"" value="""&ObjRs("adres")&""" /></div>"&_
	"<div><div>Açýklama </div>&nbsp;<input type=""text"" name=""linkAciklama"" class=""TextField"" value="""&ObjRs("aciklama")&""" /></div>"&_
	"<br /><input type=""hidden"" name=""linkID"" value="""&ObjRs("id")&""" /><input name=""LinkDegistir"" type=""submit"" value=""Kaydet"" />"&_
	"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name=""iptal"" type=""button"" value=""Ýptal"" onclick=""history.back()"" /></form>"
	Response.Write "</div>"
	ObjRs.close:set ObjRs=nothing:bottom:response.End()
	
elseif SayfaAdres(2)="Sil" then
	ObjCon.execute("DELETE from linkler where id="&SayfaAdres(3))
		
end if
end if
Response.Write "<div id=""Link-Ekle"" class=""Mesaj"">"
Response.Write "<form action="""&SiteAdres&"/islem.asp"" method=""post"" id=""linkform""><h2>Yeni Link Ekle</h2><br /><br />"&_
"<div><div>Baþlýk (*) </div><input type=""text"" name=""linkBaslik"" class=""TextField"" /></div>"&_
"<div><div>Adres (*) </div><input type=""text"" name=""linkAdres"" class=""TextField"" /></div>"&_
"<div><div>Açýklama </div><input type=""text"" name=""linkAciklama"" class=""TextField"" /></div>"&_
"<br /><input name=""LinkEkle"" type=""submit"" value=""Ekle"" /></form>"
Response.Write "</div>"
StrSql="select * from linkler order by baslik asc":ObjRs.Open StrSql,ObjCon,1,3:ToplamLink=ObjRs.recordcount
for i=1 to ToplamLink
Response.Write "<div id=""Linkler-"&ObjRs("id")&""" class=""Mesaj"" >"
Response.Write ObjRs("baslik")&"<font class=""floatRight""><a href="""&SiteAdres&"/TumLinkler/Duzenle/"&ObjRs("id")&""">Düzenle</a> - <a href="""&SiteAdres&"/TumLinkler/Sil/"&ObjRs("id")&""">Sil</a></font><br /><a href="""&ObjRs("adres")&""" target=""_blank"">"&ObjRs("adres")&"</a><br /><br />"&ObjRs("aciklama")&""
Response.Write "</div>":ObjRs.Movenext
next:ObjRs.close:set ObjRs=nothing
call bottom%>