<!-- #Include File ="tasarim.asp" --><%top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/"):Guvenlik(SayfaAdres(2))
If Ubound(SayfaAdres)=4 then page=SayfaAdres(4) else page=1
Set ObjRs=Server.CreateObject("ADODB.RecordSet"):ObjRs.open "select * from blog where gorunur=true and kategori="& SayfaAdres(2)&" order by tarih desc;",ObjCon,1,3
If ObjRs.Eof or ObjRs.Bof then
	Response.Write "<div class=""error"">aramaya çalýþtýðýnýz veri silinmiþ veya anlýk bir sorun oluþmuþ olabilir.</div>"	
else
	ObjRs.pagesize = SiteSayfalama:ObjRs.absolutepage = page:sayfa = ObjRs.pagecount
	for i=1 to ObjRs.pagesize
		if ObjRs.eof or ObjRs.bof then exit for
		Call BlogYaz(ObjRs("id"),false)'blog yazma prosedürü
		ObjRs.MoveNext
	Next
	If Not ObjRs.recordcount<SiteSayfalama+1 then
		Response.Write "<div class=""Sayfalar"">"
		If page="" or cint(page)=1 then Response.Write "<span class=""Pasif"">« ilk</span><span class=""Pasif"">&lt; geri</span>" Else	Response.Write "<a href="""&SiteAdres&"/kategori/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/1"">« ilk</a><a href="""&SiteAdres&"/kategori/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/" &  page-1 & """>&lt; geri</a>"
		'**** TANIMLAMALAR ****
		If cint(page)=2 then GeriKac=1 Else GeriKac=2
		If cint(page)=1 and sayfa<3 then IleriKac=2 Else IleriKac=3
		If cint(page)=sayfa-1 then IleriKac2=1 Else IleriKac2=2
		If cint(page)=1 then
		for y=1 to IleriKac
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/kategori/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/"&y&""">"&y&"</a>"
			next							
		ElseIf cint(page) = sayfa then
			for y=page-GeriKac to sayfa
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/kategori/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/"&y&""">"&y&"</a>"
			next
		Else
			for y=page-GeriKac to page+IleriKac2
				if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/kategori/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/"&y&""">"&y&"</a>"
			next
		End If	 
		If cint(page) = sayfa then Response.Write "<span class=""Pasif"">ileri &gt;</span><span class=""Pasif"">son »</span>" Else Response.Write "<a href="""&SiteAdres&"/kategori/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/" &  page+1 & """>ileri &gt;</a><a href="""&SiteAdres&"/kategori/"&SayfaAdres(2)&"/"&SayfaAdres(3)&"/" &  sayfa & """>son »</a>" 
		Response.Write "</div>"
	end if
End If
	ObjRs.Close:Set ObjRs = nothing
bottom%>