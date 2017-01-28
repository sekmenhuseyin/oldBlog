<!--#include file="tasarim.asp"--><%if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then response.Redirect(SiteAdres&"/")
call top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
if ubound(SayfaAdres)=2 then Guvenlik(SayfaAdres(2)):page=SayfaAdres(2) else page=1
Set ObjRs=Server.CreateObject ("ADODB.RecordSet"):StrSql="SELECT * FROM BanaMesaj ORDER BY yeni DESC,tarih DESC,saat desc;":ObjRs.Open StrSql,ObjCon,1,3
If not ObjRs.Eof or not ObjRs.Bof then
ObjRs.pagesize=SiteSayfalama*2:ObjRs.absolutepage=page:sayfa=ObjRs.pagecount
for i=1 to ObjRs.pagesize
	if ObjRs.eof or ObjRs.bof then exit for
	If Len(ObjRs("mail"))<1 then ekleyen=ObjRs("ad") Else ekleyen="<a href=""mailto:"&ObjRs("mail")&""">"&ObjRs("ad")&"</a>"
	if ObjRs("yeni")=true then OnayDugmeleri2="Okundu":OnayDugmeleri3=false else OnayDugmeleri2="Okunmadý olarak iþaretle":OnayDugmeleri3=true
	OnayDugmeleri="<a href="""&SiteAdres&"/islem.asp?MesajOkumak="&ObjRs("id")&"&islem="&OnayDugmeleri3&""">"&OnayDugmeleri2&"</a> - <a href="""&SiteAdres&"/islem.asp?MesajSil="&ObjRs("id")&""">Sil</a>"
	Response.Write "<div id=""Mesajlar-"&ObjRs("id")&""" class=""Mesaj"">"
	Response.Write "<font class=""blogTarih"">"&ekleyen&" "&TarihGoster(ObjRs("tarih"),ObjRs("saat"))&"<font class=""floatRight"">"&OnayDugmeleri&"</font></font>"
	Response.Write "<div class=""sep""></div>"&Temizle(ObjRs("ileti"))&"</div>":ObjRs.Movenext
Next
end if
'sayfa altýndaki sayfa numaralarý
If Not ObjRs.recordcount<SiteSayfalama*2+1 then
	Response.Write "<div class=""Sayfalar"">"
	If cint(page)="" or cint(page)="1" then Response.Write "<span class=""Pasif"">« ilk</span><span class=""Pasif"">&lt; geri</span>" Else Response.Write "<a href="""&SiteAdres&"/TumMesajlar/1"">« ilk</a><a href="""&SiteAdres&"/TumMesajlar/" &  page-1 & """>&lt; geri</a>" 
	'**** TANIMLAMALAR ****
	If cint(page)=2 then GeriKac=1 Else GeriKac=2
	If cint(page)=1 and sayfa < 3 then IleriKac=2 Else IleriKac=3
	If cint(page)=sayfa - 1 then IleriKac2=1 Else IleriKac2=2
	If cint(page)=1 then
		for y=1 to IleriKac
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/TumMesajlar/"&y&""">"&y&"</a>"
		next							
	ElseIf cint(page)=sayfa then
		for y=page-GeriKac to sayfa
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/TumMesajlar/"&y&""">"&y&"</a>"
		next
	Else
		for y=page-GeriKac to page+IleriKac2
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/TumMesajlar/"&y&""">"&y&"</a>"
		next
	End If	 
	If cint(page)=sayfa then Response.Write "<span class=""Pasif"">ileri &gt;</span><span class=""Pasif"">son »</span>" Else Response.Write "<a href="""&SiteAdres&"/TumMesajlar/" &  page+1 & """>ileri &gt;</a><a href="""&SiteAdres&"/TumMesajlar/" &  sayfa & """>son »</a>" 
	Response.Write "</div>"
End If
ObjRs.Close:Set ObjRs=nothing
call bottom%>