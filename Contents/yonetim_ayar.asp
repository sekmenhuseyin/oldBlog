<!--#include file="tasarim.asp"--><%if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then response.Redirect(SiteAdres&"/")
call top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
If ubound(SayfaAdres)>1 then
	if SayfaAdres(2)="Tamam" then
		response.Write "<div class=""success"">Ayarlarýnýz baþarýyla kaydedildi.</div>"
	else
		Response.Write "<div class=""error"">Ayarlarda eksik bilgi var.</div>"
	end if
end if:Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from SiteAyar":ObjRs.Open StrSql,ObjCon,1,3
response.Write("<div id=""SiteAyar"" class=""Mesaj""><form class=""horizontalForm"" method=""post"" action="""&SiteAdres&"/islem.asp""><h2>Site Ayarlarý</h2><br /><br />")
for i=0 to 10
response.Write("<label>"&ObjRs(i).name&":&nbsp;</label><input type=""text"" name="""&ObjRs(i).name&""" class=""TextField"" value="""&ObjRs(i)&""" />")
next:ObjRs.close:set ObjRs=nothing
response.Write("<br /><label>&nbsp;</label><input name=""SiteAyarKaydet"" type=""submit"" value=""Güncelle"" /></form><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /></div>")
call bottom%>