<!-- #Include File ="tasarim.asp" --><%call top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
if ubound(SayfaAdres)>1 then txtAra=SayfaAdres(2):txtAra=right(txtAra,len(txtAra)-instr(txtAra,"=")):txtAra=TurkceHarfler(txtAra)
Response.Write "<div id=""Arama"" class=""Mesaj""><h2>Arama Sonuçlarý</h2><br/><b>&#34;"&txtAra&"&#34;</b> için bulunan sonuçlar gösteriliyor.</div>"
Response.Write "<div id=""Sonuc"" class=""Mesaj"">"
Set ObjRs = Server.CreateObject ("ADODB.RecordSet"):StrSql = " select * from blog where gorunur=true and (baslik like '%"&txtAra&"%' or metin like '%"&txtAra&"%') order by baslik asc;":ObjRs.Open StrSql, ObjCon, 1, 3
	If ObjRs.eof then
		Response.Write "Maalesef aradýðýnýz kelime ile eþleþen bir sonuç bulunamadý."
	Else
		for i=1 to 20
		If ObjRs.eof then exit for
			Response.Write "<a href="""&SiteAdres&"/blog/"&ObjRs("id")&"/"&SEO_Olustur(ObjRs("baslik"))&".html"">"&ObjRs("baslik")&"</a><br/>"
			ObjRs.movenext	
		Next
	End If
ObjRs.Close:Set ObjRs = nothing:Response.Write "</div>"
call bottom%>