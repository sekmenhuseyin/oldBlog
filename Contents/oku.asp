<!--#include file="tasarim.asp"--><%SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/"):Guvenlik(SayfaAdres(2))
Set ObjRs=Server.CreateObject ("ADODB.RecordSet"):ObjRs.open "select * from blog where id="&SayfaAdres(2)&";", ObjCon, 1, 3
if not ObjRs.eof or not ObjRs.bof then
KonuBaslik=ObjRs("baslik")&"&nbsp;-&nbsp;":top
if ubound(SayfaAdres)=4 then
if SayfaAdres(4)="yorum-ok" then
	yorumMSG="<div class=""success"">Yorumunuz eklenmiþtir.</div>"
elseif SayfaAdres(4)="yorum-hata" then
	yorumMSG="<div class=""error"">Boþ alanlar var.</div>"
elseif SayfaAdres(4)="Yeni" or SayfaAdres(4)="Kaydet" then
	response.Write "<div class=""success"">Blogunuz baþarýyla kaydedildi.</div>"
else
	yorumMSG=""
end if
end if
if request.Cookies("aBlog")(""&ObjRs("id")&"")="" and ObjRs("gorunur")=true and Session("Yonetici")<>true then ObjRs("okunma")=ObjRs("okunma")+1:ObjRs.update:Response.Cookies("aBlog")(""&ObjRs("id")&"")="okundu":Response.Cookies("aBlog").expires=SiteTarih+1:Response.Cookies("aBlog").Path=""
if Session("Yonetici")=true or ObjRs("gorunur")=true then Call BlogYaz(ObjRs("id"),true) else ObjRs.Close:Set ObjRs=nothing:response.Redirect(SiteAdres&"/SayfaBulunamadi")
else
response.Redirect(SiteAdres&"/SayfaBulunamadi")
end if
ObjRs.Close:Set ObjRs=nothing
bottom%>