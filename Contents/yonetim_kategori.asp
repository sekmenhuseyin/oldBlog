<!--#include file="tasarim.asp"--><%if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then response.Redirect(SiteAdres&"/")
call top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
if ubound(SayfaAdres)=2 then
	if SayfaAdres(2)="Yeni" or  SayfaAdres(2)="Kaydet" then
		response.Write "<div class=""success"">Kategoriniz baþarýyla kaydedildi.</div>"
	else
		Response.Write "<div class=""error"">Kategorinizde eksik bilgi var</div>"
	end if
end if:Set ObjRs=Server.CreateObject("ADODB.RecordSet")
if ubound(SayfaAdres)>2 then
Guvenlik(SayfaAdres(3))
if SayfaAdres(2)="Duzenle" then
	StrSql="select * from kategoriler where id="&SayfaAdres(3):ObjRs.Open StrSql,ObjCon,1,3
	If ObjRs.Eof or ObjRs.Bof then
		Response.Write "<div class=""error"">aramaya çalýþtýðýnýz veri silinmiþ veya anlýk bir sorun oluþmuþ olabilir.</div>"	
	else
		Response.Write "<div id=""Kategori-Guncelle"" class=""Mesaj"">"
		Response.Write "<form class=""horizontalForm"" action="""&SiteAdres&"/islem.asp"" method=""post"" id=""Kategoriform""><h2>Kategori Deðiþtir</h2><br /><br />"&_
		"<div>Baþlýk (*) </div>&nbsp;<input type=""text"" name=""KategoriBaslik"" class=""TextField"" value="""&ObjRs("adi")&""" /><br />"&_
		"<br /><input type=""hidden"" name=""KategoriID"" value="""&ObjRs("id")&""" /><input name=""KategoriDegistir"" type=""submit"" value=""Kaydet"" />"&_
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name=""iptal"" type=""button"" value=""Ýptal"" onclick=""history.back()"" /></form>"
		Response.Write "</div>"
	end if:ObjRs.close:set ObjRs=nothing:bottom:response.End()
	
elseif SayfaAdres(2)="Sil" then
	ObjCon.execute("DELETE from kategoriler where id="&SayfaAdres(3))
	
end if
end if

Response.Write "<div id=""Kategori-Ekle"" class=""Mesaj"">"
Response.Write "<form action="""&SiteAdres&"/islem.asp"" method=""post"" id=""Kategoriform""><h2>Yeni Kategori Ekle</h2><br /><br />"&_
"<div>Baþlýk (*) </div>&nbsp;<input type=""text"" name=""KategoriBaslik"" class=""TextField"" /><br />"&_
"<br /><input name=""KategoriEkle"" type=""submit"" value=""Ekle"" /></form>"
Response.Write "</div>"
StrSql="SELECT * FROM kategoriler ORDER BY adi;":ObjRs.Open StrSql,ObjCon,1,3:ToplamLink=ObjRs.recordcount
for i=1 to ToplamLink
Response.Write "<div id=""Linkler-"&ObjRs("id")&""" class=""Mesaj"" >"
Response.Write ObjRs("adi")&"<font class=""blogTarih floatRight""><a href="""&SiteAdres&"/TumKategoriler/Duzenle/"&ObjRs("id")&""">Düzenle</a> - <a href="""&SiteAdres&"/TumKategoriler/Sil/"&ObjRs("id")&""">Sil</a></font>"
Response.Write "</div>":ObjRs.Movenext
next:ObjRs.close:set ObjRs=nothing
call bottom%>