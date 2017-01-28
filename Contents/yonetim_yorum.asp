<!--#include file="tasarim.asp"--><%if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then response.Redirect(SiteAdres&"/")
call top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
if ubound(SayfaAdres)=2 then Guvenlik(SayfaAdres(2)):page=SayfaAdres(2) else page=1
Set ObjRsYorum=Server.CreateObject ("ADODB.RecordSet"):StrSql="SELECT Blog.baslik,yorum.* FROM yorum INNER JOIN Blog ON yorum.blog=Blog.id ORDER BY yorum.onay DESC , yorum.tarih DESC;":ObjRsYorum.Open StrSql,ObjCon,1,3
If not ObjRsYorum.Eof or not ObjRsYorum.Bof then
ObjRsYorum.pagesize=SiteSayfalama*2:ObjRsYorum.absolutepage=page:sayfa=ObjRsYorum.pagecount
for i=1 to ObjRsYorum.pagesize
	if ObjRsYorum.eof or ObjRsYorum.bof then exit for
	if ObjRsYorum("onay")=false then OnayDugmeleri="<a href="""&SiteAdres&"/islem.asp?YorumOnayla="&ObjRsYorum("id")&""">Onayla</a>" else OnayDugmeleri="<a href="""&SiteAdres&"/blog/"&ObjRsYorum("blog")&"/"&SEO_Olustur(ObjRsYorum("baslik"))&".html#Mesaj-Yorumlar"">Cevap Yaz</a> - <a href="""&SiteAdres&"/islem.asp?YorumOnaylamama="&ObjRsYorum("id")&""">Sakla</a>"
	OnayDugmeleri=OnayDugmeleri&" - <a href="""&SiteAdres&"/islem.asp?YorumSil="&ObjRsYorum("id")&""">Sil</a>"
	Response.Write "<div id=""Yorumlar-"&ObjRsYorum("id")&""" class=""Mesaj""><a href="""&SiteAdres&"/blog/"&ObjRsYorum("blog")&"/"&SEO_Olustur(ObjRsYorum("baslik"))&".html"">"&ObjRsYorum("baslik")&"</a><font class=""floatRight"">"&OnayDugmeleri&"</font><div class=""sep""></div>"
	If Len(ObjRsYorum("web")) < 1 then ekleyen=ObjRsYorum("ekleyen") Else ekleyen="<a href=""http://"&ObjRsYorum("web")&""" target=""_blank"">"&ObjRsYorum("ekleyen")&"</a>"
	Response.Write "<font class=""blogTarih"">"&ekleyen&"<font class=""floatRight"">"&TarihGoster(ObjRsYorum("tarih"),ObjRsYorum("saat"))&"</font></font>"
	Response.Write "<div class=""sep""></div>"&Temizle(ObjRsYorum("metin"))&"</div>":ObjRsYorum.Movenext
Next
end if
If Not ObjRsYorum.recordcount<SiteSayfalama*2+1 then
	Response.Write "<div class=""Sayfalar"">"
	If cint(page)="" or cint(page)="1" then Response.Write "<span class=""Pasif"">« ilk</span><span class=""Pasif"">&lt; geri</span>" Else Response.Write "<a href="""&SiteAdres&"/TumYorumlar/1"">« ilk</a><a href="""&SiteAdres&"/TumYorumlar/" &  page-1 & """>&lt; geri</a>" 
	'**** TANIMLAMALAR ****
	If cint(page)=2 then GeriKac=1 Else GeriKac=2
	If cint(page)=1 and sayfa < 3 then IleriKac=2 Else IleriKac=3
	If cint(page)=sayfa - 1 then IleriKac2=1 Else IleriKac2=2
	If cint(page)=1 then
		for y=1 to IleriKac
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/TumYorumlar/"&y&""">"&y&"</a>"
		next							
	ElseIf cint(page)=sayfa then
		for y=page-GeriKac to sayfa
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/TumYorumlar/"&y&""">"&y&"</a>"
		next
	Else
		for y=page-GeriKac to page+IleriKac2
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/TumYorumlar/"&y&""">"&y&"</a>"
		next
	End If	 
	If cint(page)=sayfa then Response.Write "<span class=""Pasif"">ileri &gt;</span><span class=""Pasif"">son »</span>" Else Response.Write "<a href="""&SiteAdres&"/TumYorumlar/" &  page+1 & """>ileri &gt;</a><a href="""&SiteAdres&"/TumYorumlar/" &  sayfa & """>son »</a>" 
	Response.Write "</div>"
End If
ObjRsYorum.Close:Set ObjRsYorum=nothing
call bottom%>