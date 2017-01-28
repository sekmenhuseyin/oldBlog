<!--#include file="tasarim.asp"--><%if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then response.Redirect(SiteAdres&"/")
call top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
if ubound(SayfaAdres)=2 then
	if SayfaAdres(2)="Tamam" then
		response.Write "<div class=""success"">Etiketiniz baþarýyla deðiþtirildi.</div>"
	else
		Response.Write "<div class=""error"">Eksik bilgi var.</div>"
	end if
end if:Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="SELECT * FROM etiket_ad ORDER BY etiket asc;":ObjRs.Open StrSql,ObjCon,1,3:ObjRsCount=ObjRs.recordcount
response.Write("<div id=""SiteAyar"" class=""Mesaj""><h2>Etiketler</h2><br /><br />")
for i=0 to ObjRsCount-1
response.Write("<div><form class=""horizontalForm"" method=""post"" action="""&SiteAdres&"/islem.asp"">"&_
"<input type=""text"" name=""etiket"" class=""TextField"" value="""&ObjRs(1)&""" />"&_
" <a href="""&SiteAdres&"/islem.asp?EtiketSil="&ObjRs(0)&""">Sil</a> "&_
"<input type=""hidden"" name=""id"" value="""&ObjRs(0)&""" /><input name=""EtiketKaydet"" type=""submit"" value=""Güncelle"" />"&_
"</form><br /></div>")
ObjRs.movenext
next:ObjRs.close:set ObjRs=nothing:response.Write("</div>")
call bottom%>