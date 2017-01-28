<!-- #Include File ="tasarim.asp" --><%top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/"):Guvenlik(SayfaAdres(2))
Set ObjRs = Server.CreateObject ("ADODB.RecordSet"):ObjRs.open "SELECT etiket_bulut.*, Blog.tarih FROM etiket_bulut INNER JOIN Blog ON etiket_bulut.blog_id=Blog.id where blog.gorunur=true and etiket_id="&SayfaAdres(2)&" ORDER BY Blog.tarih DESC;", ObjCon, 1, 3
If ObjRs.Eof or ObjRs.Bof then
	Response.Write "<div class=""error"">aramaya çalýþtýðýnýz veri silinmiþ veya anlýk bir sorun oluþmuþ olabilir.</div>"	
else
	If Ubound(SayfaAdres)=4 then page=SayfaAdres(4) Else page=1
	ObjRs.pagesize=SiteSayfalama:ObjRs.absolutepage=page:sayfa=ObjRs.pagecount
	for i=1 to ObjRs.pagesize
		if ObjRs.eof or ObjRs.bof then exit for
		Call BlogYaz(ObjRs("blog_id"),false)'blog yazma prosedr
		ObjRs.MoveNext
	Next
end if
If Not ObjRs.recordcount<SiteSayfalama+1 then
	Response.Write "<div class=""Sayfalar"">"
	If cint(page) = "" or cint(page) = "1" then Response.Write "<span class=""Pasif""> ilk</span><span class=""Pasif"">&lt; geri</span>" Else Response.Write "<a href="""&SiteAdres&"/etiket/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/1""> ilk</a><a href="""&SiteAdres&"/etiket/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/" &  page-1 & """>&lt; geri</a>"
	'**** TANIMLAMALAR ****
	If cint(page)=2 then GeriKac=1 Else GeriKac=2
	If cint(page)=1 and ObjRs.pagecount<3 then IleriKac=2 Else IleriKac=3
	If cint(page)=ObjRs.pagecount-1 then IleriKac2=1 Else IleriKac2=2
	If cint(page)=1 then
	for y=1 to IleriKac
		if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/etiket/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/"&y&""">"&y&"</a>"
		next							
	ElseIf cint(page) = sayfa then
		for y=page-GeriKac to sayfa
		if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/etiket/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/"&y&""">"&y&"</a>"
		next
	Else
		for y=page-GeriKac to page+IleriKac2
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/etiket/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/"&y&""">"&y&"</a>"
		next
	End If	 
	If cint(page) = sayfa then Response.Write "<span class=""Pasif"">ileri &gt;</span><span class=""Pasif"">son </span>" Else Response.Write "<a href="""&SiteAdres&"/etiket/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/" &  page+1 & """>ileri &gt;</a><a href="""&SiteAdres&"/etiket/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/" &  sayfa & """>son </a>" 
	Response.Write "</div>"
End If
ObjRs.Close:Set ObjRs = nothing
bottom%>