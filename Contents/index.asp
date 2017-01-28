<!-- #Include File="tasarim.asp" --><%top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):page=1':SiteSayfalama=1
if SayfaAdres>"" then SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/"):if SayfaAdres(1)>"" then Guvenlik(SayfaAdres(1)):page=SayfaAdres(1)
if Session("Yonetici")=true then sql="select * from blog order by gorunur desc,tarih desc,saat desc;" else sql="select * from blog where gorunur=true order by tarih desc,saat desc;"
Set ObjRs=Server.CreateObject ("ADODB.RecordSet"):ObjRs.open sql, ObjCon, 1, 3
if not ObjRs.eof or not ObjRs.bof then
	ObjRs.pagesize=SiteSayfalama:sayfa=ObjRs.pagecount:if cint(page)>cint(sayfa) then page=sayfa
	ObjRs.absolutepage=page:for i=1 to ObjRs.pagesize
		if ObjRs.eof or ObjRs.bof then exit for
		Call BlogYaz(ObjRs("id"),false):ObjRs.MoveNext'iþte burada blog yazma prosedürü çaðrýlýyor...
	Next
end if
If Not ObjRs.recordcount<SiteSayfalama+1 then'burdada sayfanýn altýndaki sayfa numaralarý var.
	Response.Write "<div class=""Sayfalar"">"
	If cint(page)=sayfa then Response.Write "<span class=""Pasif"">&lt;&lt; Ýlk Yazýlarým</span><span class=""Pasif"">&lt; Eski Yazýlarým</span>" Else Response.Write "<a href="""&SiteAdres&"/" &  sayfa & """>&lt;&lt; Ýlk Yazýlarým</a><a href="""&SiteAdres&"/" &  page+1 & """>&lt; Eski Yazýlarým</a>" 
	'**** TANIMLAMALAR ****
	If cint(page)=2 then GeriKac=1 Else GeriKac=2
	If cint(page)=1 and ObjRs.pagecount < 3 then IleriKac=2 Else IleriKac=3
	If cint(page)=ObjRs.pagecount - 1 then IleriKac2=1 Else IleriKac2=2
	If cint(page)=1 then
		for y=IleriKac to 1 step -1
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/"&y&""">"&y&"</a>"
		next							
	ElseIf cint(page)=sayfa then
		for y=sayfa to page-GeriKac step -1
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/"&y&""">"&y&"</a>"
		next
	Else
		for y=page+IleriKac2 to page-GeriKac step -1
			if y=cint(page) then response.write "<span class=""Ayni"">"&y&"</span>" else response.write "<a href="""&SiteAdres&"/"&y&""">"&y&"</a>"
		next
	End If	 
	If cint(page)="" or cint(page)="1" then Response.Write "<span class=""Pasif"">Yeni Yazýlarým &gt;</span><span class=""Pasif"">Son Yazýlarým &gt;&gt;</span>" Else Response.Write "<a href="""&SiteAdres&"/" &  page-1 & """>Yeni Yazýlarým &gt;</a><a href="""&SiteAdres&"/"">Son Yazýlarým &gt;&gt;</a>" 
	Response.Write "</div>"
End If
ObjRs.Close:Set ObjRs=nothing
bottom%>