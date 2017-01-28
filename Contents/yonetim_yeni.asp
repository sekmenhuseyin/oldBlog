<!--#include file="tasarim.asp"--><%if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then response.Redirect(SiteAdres&"/")
call top:SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
if ubound(SayfaAdres)>1 then
	if isnumeric(SayfaAdres(2))=true then
		Set ObjRs=Server.CreateObject ("ADODB.RecordSet"):ObjRs.open "select * from blog where id="&SayfaAdres(2)&";", ObjCon, 1, 3
		BlogID=ObjRs("id"):BlogBaslik=ObjRs("baslik"):BlogMetin=ObjRs("metin"):BlogKategori=ObjRs("kategori"):BlogEtkin=ObjRs("gorunur")
		ObjRs.Close:Set ObjRs=nothing
	else
		Response.Write "<div class=""error"">BLOG için baþlýk ve içerik lazým</div>"
	end if
end if%><div id="Mesaj-Oku" class="Mesaj"><form method="post" action="<%=SiteAdres%>/islem.asp">
<h2><%if BlogID="" then response.Write("Yeni Yazý Ekle") else response.Write("Yazý Düzenle")%></h2><br /><br />
<div><label for="ad">Yazýnýn Adý: </label><input type="text" name="ad" class="TextField" value="<%=BlogBaslik%>" /></div>
<div>Ýçerik:<br />
<textarea name="icerik" rows="30" cols="20" class="TextArea"><%=BlogMetin%></textarea></div>
<div><br />Kategori: <select name="kategori" class="TextField"><option value=""></option><%Set ObjRsKat=Server.CreateObject ("ADODB.RecordSet"):StrSql="select * from kategoriler order by adi asc":ObjRsKat.Open StrSql,ObjCon,1,3:for i=1 to ObjRsKat.recordcount:response.write("<option value="""&ObjRsKat("id")&"""")
if BlogKategori=ObjRsKat("id") then response.Write(" selected")
response.Write(">"&ObjRsKat("adi")&"</option>"):ObjRsKat.MoveNext:Next:ObjRsKat.close
if BlogID>"" then'eðer düzenlemek için bir id gönderilmiþse o blogun etiketlerinin getir....
StrSql="SELECT etiket_ad.etiket, etiket_bulut.blog_id FROM etiket_bulut INNER JOIN etiket_ad ON etiket_bulut.etiket_id=etiket_ad.id GROUP BY etiket_ad.etiket, etiket_bulut.blog_id HAVING etiket_bulut.blog_id="&BlogID&" order by etiket_ad.etiket desc;":ObjRsKat.Open StrSql,ObjCon,1,3:ObjRsKatCount=ObjRsKat.recordcount
if ObjRsKatCount>0 then
for i=0 to ObjRsKatCount-1
BlogEtiketler=ObjRsKat(0)&", "&BlogEtiketler:ObjRsKat.movenext
next:BlogEtiketler=left(BlogEtiketler,len(BlogEtiketler)-2)
end if:ObjRsKat.close
end if%></select></div>
<div>Etiketler: <input type="text" name="etiket" id="etiket" class="TextField" value="<%=BlogEtiketler%>" /> <font class="sayfaTelif">Etiketleri virgül ile ayýrýn.</font></div><br /><input type="hidden" name="id" value="<%=BlogID%>" /><input name="BlogKaydet" type="submit" value="Kaydet" />&nbsp;&nbsp;&nbsp;
<input name="BlogYayýnla" type="submit" value="Yayýnla" />&nbsp;&nbsp;&nbsp;<input name="iptal" type="button" value="Ýptal" onclick="history.back()" /></form><%
'etiket array for auto complete
StrSql="SELECT etiket FROM etiket_ad ORDER BY etiket;":ObjRsKat.Open StrSql,ObjCon,1,3:ObjRsKatCount=ObjRsKat.recordcount:BlogEtiketler=""
for i=0 to ObjRsKatCount-1
BlogEtiketler="'"&ObjRsKat(0)&"',"&BlogEtiketler:ObjRsKat.movenext
next:BlogEtiketler=left(BlogEtiketler,len(BlogEtiketler)-1):ObjRsKat.close:Set ObjRsKat=nothing
%></form></div><script type="text/javascript" src="<%=SiteAdres%>/Contents/Editor/ckeditor.js"></script><script language="javascript" type="text/javascript" src="<%=SiteAdres%>/Contents/include/actb.js"></script><script language="javascript" type="text/javascript" src="<%=SiteAdres%>/Contents/include/common.js"></script><%
if SayfaAdres(1)<>"TemizDuzenle" then%><script type="text/javascript">CKEDITOR.replace('icerik');customarray=new Array(<%=BlogEtiketler%>);actb(document.getElementById('etiket'),customarray);</script><%end if
call bottom%>