<!-- #Include File="Contents/abt-yapilandirma.asp"--><!-- #Include File="Contents/abt-fonksiyonlar.asp"--><meta http-equiv="Content-Language" content="tr"><%
response.Write("<h1>Ýþleminiz sürmektedir, lütfen bekleyiniz.</h1><h2>Çok uzun sürerse tarayýcýnýzý yenileyin.</h2>")
if request.Form("LoginSubmit")<>"" then'login iþlemleri
	set rst=Server.CreateObject("ADODB.RecordSet"):rst.open "select * from SiteAyar",ObjCon,3,3
	if QS_Temizle(Request.Form("Lakap"))=rst("Lakap") and QS_Temizle(Request.Form("Parola"))=rst("Parola") then Session("Yonetici")=true:Session.Timeout=20
	rst.close:Response.Redirect Request.ServerVariables("HTTP_REFERER")
	
elseif request.Form("BanaMesajGonder")<>"" then'bana mesaj gönderme iþlemleri
	ekleyen=QS_Temizle(request.Form("ekleyen")):eposta=QS_Temizle(request.Form("eposta")):metin=QS_Temizle(request.Form("metin")):RegSecurityCode=QS_Temizle(request.Form("guvenlik"))
	if ekleyen="" or eposta="" or metin="" or CStr(RegSecurityCode)<>CStr(Session("CAPTCHA")) then
		response.Redirect(SiteAdres&"/iletisim/hata")
	else
		Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from BanaMesaj":ObjRs.Open StrSql,ObjCon,1,3
		ObjRs.addnew:ObjRs("ad")=ekleyen:ObjRs("mail")=eposta:ObjRs("ileti")=metin:ObjRs.update:ObjRs.Close:set ObjRs=nothing
		Response.Cookies("aBlog")("Ekleyen")=ekleyen:Response.Cookies("aBlog")("Eposta")=eposta:Response.Cookies("aBlog").expires=SiteTarih+360:Response.Cookies("aBlog").Path=""
		'Call BanaMailAtHaberimOlsun()		
		response.Redirect(SiteAdres&"/iletisim/Tamam")
	end if

elseif request.Form("EtiketKaydet")<>"" then'etiket adý deðiþtirme iþlemlerim
	EtiketID=QS_Temizle(request.Form("id")):EtiketAd=QS_Temizle(request.Form("etiket"))
	if EtiketID="" or EtiketAd="" then response.Redirect(SiteAdres&"/TumEtiketler/EksikBilgi")
	Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from etiket_ad where id="&trim(EtiketID):ObjRs.Open StrSql,ObjCon,1,3
	if not ObjRs.eof then ObjRs("etiket")=trim(EtiketID):ObjRs.update
	ObjRs.close:Set ObjRs=Nothing:response.Redirect(SiteAdres&"/TumEtiketler/Tamam")

elseif request.QueryString("EtiketSil")<>"" then'etiket adý deðiþtirme iþlemlerim
	if isnumeric(QS_Temizle(request.QueryString("EtiketSil")))=true then ObjCon.execute("DELETE from etiket_ad where id="&QS_Temizle(request.QueryString("EtiketSil")))
	response.Redirect(SiteAdres&"/TumEtiketler/Tamam")

elseif request.Form("SiteAyarKaydet")<>"" then'sitenin ayarlarýný kaydederken
	for each fn in request.Form
		if trim(request.Form(fn))="" then response.Redirect(SiteAdres&"/Ayarlar/EksikBilgi")
	next
	Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from SiteAyar":ObjRs.Open StrSql,ObjCon,1,3
	for each fn in request.Form
		if fn<>"SiteAyarKaydet" then ObjRs(fn)=trim(request.Form(fn))
	next
	ObjRs.update:ObjRs.close:set ObjRs=nothing:response.Redirect(SiteAdres&"/Ayarlar/Tamam")	

elseif request.Form("BlogKaydet")<>"" then'bir blogu düzenledikten sonra kaydediþim...
	BlogID=QS_Temizle(request.Form("id")):BlogAd=QS_Temizle(request.Form("ad")):BlogIcerik=request.Form("icerik"):BlogKategori=QS_Temizle(request.Form("kategori")):BlogEtiket=QS_Temizle(request.Form("etiket")):if BlogKategori="" then BlogKategori=0
	if BlogID="" then Call NewBlog(false) else Call SaveBlog

elseif request.Form("Yayýnla")<>"" then'yeni bir blog yazýlýyor ve kaydediliyor...
	BlogAd=QS_Temizle(request.Form("ad")):BlogIcerik=request.Form("icerik"):BlogKategori=QS_Temizle(request.Form("kategori")):BlogEtiket=QS_Temizle(request.Form("etiket")):if BlogKategori="" then BlogKategori=0
	if BlogID="" then Call NewBlog(true) else Call SaveBlog

elseif request.Form("LinkEkle")<>"" then'yeni bir link ekleniyor
	linkBaslik=QS_Temizle(request.Form("linkBaslik")):linkAdres=QS_Temizle(request.Form("linkAdres")):linkAciklama=QS_Temizle(request.Form("linkAciklama"))
	if linkBaslik="" then
		response.Redirect(SiteAdres&"/TumLinkler/Ad")
	elseif linkAdres="" then
		response.Redirect(SiteAdres&"/TumLinkler/Adres")
	else
		Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from linkler":ObjRs.Open StrSql,ObjCon,1,3
		ObjRs.addnew:ObjRs("baslik")=linkBaslik:ObjRs("adres")=linkAdres:ObjRs("aciklama")=linkAciklama:ObjRs.update:ObjRs.close:Set ObjRs=Nothing
		response.Redirect(SiteAdres&"/TumLinkler/Yeni")
	end if
	
elseif request.Form("LinkDegistir")<>"" then'var olan link deðiþtiriliyor
	linkID=QS_Temizle(request.Form("linkID")):linkBaslik=QS_Temizle(request.Form("linkBaslik")):linkAdres=QS_Temizle(request.Form("linkAdres")):linkAciklama=QS_Temizle(request.Form("linkAciklama"))
	if linkBaslik="" then
		response.Redirect(SiteAdres&"/TumLinkler/Ad")
	elseif linkAdres="" then
		response.Redirect(SiteAdres&"/TumLinkler/Adres")
	else
		Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from linkler where id="&linkID:ObjRs.Open StrSql,ObjCon,1,3
		ObjRs("baslik")=linkBaslik:ObjRs("adres")=linkAdres:ObjRs("aciklama")=linkAciklama:ObjRs.update:ObjRs.close:Set ObjRs=Nothing
		response.Redirect(SiteAdres&"/TumLinkler/Kaydet")
	end if

elseif request.Form("KategoriEkle")<>"" then'yeni bir kategori ekleniyor
	KategoriBaslik=QS_Temizle(request.Form("KategoriBaslik"))
	if KategoriBaslik="" then
		response.Redirect(SiteAdres&"/TumKategoriler/EksikBilgi")
	else
		Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from kategoriler":ObjRs.Open StrSql,ObjCon,1,3
		ObjRs.addnew:ObjRs("adi")=KategoriBaslik:ObjRs.update:ObjRs.close:Set ObjRs=Nothing
		response.Redirect(SiteAdres&"/TumKategoriler/Yeni")
	end if
	
elseif request.Form("KategoriDegistir")<>"" then'var olan kategori deðiþtiriliyor
	KategoriID=QS_Temizle(request.Form("KategoriID")):KategoriBaslik=QS_Temizle(request.Form("KategoriBaslik"))
	if KategoriBaslik="" or KategoriID="" then
		response.Redirect(SiteAdres&"/TumKategoriler/EksikBilgi")
	else
		Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from kategoriler where id="&KategoriID:ObjRs.Open StrSql,ObjCon,1,3
		ObjRs("adi")=KategoriBaslik:ObjRs.update:ObjRs.close:Set ObjRs=Nothing
		response.Redirect(SiteAdres&"/TumKategoriler/Kaydet")
	end if

elseif request.Form("YorumGonder")<>"" then'birisi yorum yazýyor, bu da kaydetme iþlemleri
	if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then
		txtekleyen=QS_Temizle(""&Request.Form("ekleyen")&""):txteposta=QS_Temizle(""&Request.Form("eposta")&""):txtweb=QS_Temizle(""&Request.Form("web")&"")
		if txtekleyen=AdSoyad then txtekleyen=""'eðer benim adýmým kullanýyorsa onu sil
		if txteposta="YoneticiEposta" then txteposta=""'eðer benim milimi kullanýyorsa onu da sil
	else
		txtekleyen=AdSoyad:txteposta="YoneticiEposta":txtweb=SiteAdres&"/"'ben isem bilgilerim direk yaz...
	end if:txtmetin=QS_Temizle(""&Request.Form("metin")&""):txtblog=QS_Temizle(""&Request.Form("Blog")&""):RegSecurityCode=QS_Temizle(request.Form("strCAPTCHA"))
	If len(txtekleyen)<1 or len(txteposta)<6 or len(txtmetin)<1 or len(txtblog)<1 or CStr(RegSecurityCode)<>CStr(Session("CAPTCHA")) then
		response.Redirect(SiteAdres&"/blog/"&txtblog&"/"&Request.Form("SEO")&".html/yorum-hata#Mesaj-Yorumlar")'eksik bilgi
	Else
		Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from yorum":ObjRs.Open StrSql,ObjCon,1,3
		if ObjRs("yorumlanabilir")=true then ObjRs.AddNew:ObjRs("ekleyen")=txtekleyen:ObjRs("eposta")=txteposta:ObjRs("web")=txtweb:ObjRs("metin")=txtmetin:ObjRs("blog")=txtblog:ObjRs("tarih")=SiteTarih:ObjRs.Update:ObjRs.Close:Set ObjRs=Nothing
		if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then Response.Cookies("aBlog")("Ekleyen")=txtekleyen:Response.Cookies("aBlog")("Eposta")=txteposta:Response.Cookies("aBlog")("Web")=txtweb:Response.Cookies("aBlog").expires=SiteTarih+360:Response.Cookies("aBlog").Path=""
		'Call BanaMailAtHaberimOlsun()		
		response.Redirect(SiteAdres&"/blog/"&txtblog&"/"&Request.Form("SEO")&".html/yorum-ok#Mesaj-Yorumlar")
	End If

elseif request.QueryString("MesajSil")<>"" then'bana gelen mesajlarý silme iþlemleri
	if isnumeric(request.QueryString("MesajSil"))=true then ObjCon.execute("DELETE from BanaMesaj where id="&request.QueryString("MesajSil"))
	Response.Redirect Request.ServerVariables("HTTP_REFERER")

elseif request.QueryString("MesajOkumak")<>"" then'bana gelen mesajlarý silme iþlemleri
	if isnumeric(request.QueryString("MesajOkumak"))=true then ObjCon.execute("UPDATE BanaMesaj set yeni="&request.QueryString("islem")&" where id="&request.QueryString("MesajOkumak"))
	Response.Redirect Request.ServerVariables("HTTP_REFERER")

elseif request.QueryString("YorumOnayla")<>"" then'benim yorumlarý onaylama iþlemlerim
	if isnumeric(request.QueryString("YorumOnayla"))=true then ObjCon.execute("UPDATE yorum set onay=true where id="&request.QueryString("YorumOnayla"))
	Response.Redirect Request.ServerVariables("HTTP_REFERER")

elseif request.QueryString("YorumOnaylamama")<>"" then'benim yorumlarý onaylama iþlemlerim
	if isnumeric(request.QueryString("YorumOnaylamama"))=true then ObjCon.execute("UPDATE yorum set onay=false where id="&request.QueryString("YorumOnaylamama"))
	Response.Redirect Request.ServerVariables("HTTP_REFERER")

elseif request.QueryString("YorumSil")<>"" then'benim yorumlarý silme iþlemlerim
	if isnumeric(request.QueryString("YorumSil"))=true then ObjCon.execute("DELETE from yorum where id="&request.QueryString("YorumSil"))
	Response.Redirect Request.ServerVariables("HTTP_REFERER")

elseif request.QueryString("BlogPasif")<>"" then'bir blogu gounmez hale getiriþim
	if isnumeric(request.QueryString("BlogPasif"))=true then ObjCon.execute("UPDATE blog set gorunur=false where id="&request.QueryString("BlogPasif"))
	Response.Redirect Request.ServerVariables("HTTP_REFERER")

elseif request.QueryString("BlogAktif")<>"" then'blogu tekrar gorunur hale getiriþim
	if isnumeric(request.QueryString("BlogAktif"))=true then ObjCon.execute("UPDATE blog set gorunur=true, tarih=date(), saat=time() where id="&request.QueryString("BlogAktif"))
	Response.Redirect Request.ServerVariables("HTTP_REFERER")

elseif request.QueryString("BlogSil")<>"" then've bir blogu, etiketlerinin ve yorumlarýný siliþim....
	if isnumeric(request.QueryString("BlogSil"))=true then 
		ObjCon.execute("DELETE from yorum where blog="&request.QueryString("BlogSil"))
		ObjCon.execute("DELETE from etiket_bulut where blog_id="&request.QueryString("BlogSil"))
		ObjCon.execute("DELETE from blog where id="&request.QueryString("BlogSil"))
	end if

elseif request.QueryString("Logout")<>"" then'logout iþlemleri
	Session("Yonetici")=null

end if
Response.Redirect (SiteAdres&"/")'iþlem ne olursa olsun en sonunda tekrar anasayfaya dönüyor....

Sub SaveBlog
if BlogAd="" then
	response.Redirect(SiteAdres&"/Duzenle/"&BlogID&"/Ad")
elseif BlogIcerik="" then
	response.Redirect(SiteAdres&"/Duzenle/"&BlogID&"/Icerik")
else
	Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from blog where id="&BlogID:ObjRs.Open StrSql,ObjCon,1,3
	ObjRs("baslik")=BlogAd:ObjRs("metin")=BlogIcerik:ObjRs("kategori")=BlogKategori:ObjRs.update:ObjRs.Close'blogdaki deðiþiklikler kaydedildi
	ObjCon.execute("DELETE from etiket_bulut where blog_id="&BlogID):BlogEtiket=split(BlogEtiket&",",",")'önceki etiketler siliniyor
	if ubound(BlogEtiket)>0 then'yeni etiketler yazýlýyor
		for i=0 to ubound(BlogEtiket)
			if trim(BlogEtiket(i))>"" then
			StrSql="select * from etiket_ad where etiket='"&trim(BlogEtiket(i))&"'":ObjRs.Open StrSql,ObjCon,1,3
			if ObjRs.eof then ObjRs.addnew:ObjRs("etiket")=trim(BlogEtiket(i)):ObjRs.update'etiket daha önce yosa oluþturuluyor
			EtiketID=ObjRs("id"):ObjRs.Close:StrSql="select * from etiket_bulut":ObjRs.Open StrSql,ObjCon,1,3
			ObjRs.addnew:ObjRs("blog_id")=BlogID:ObjRs("etiket_id")=EtiketID:ObjRs.update:ObjRs.close
			end if
		next		
	end if:Set ObjRs=Nothing
	response.Redirect(SiteAdres&"/blog/"&BlogID&"/"&SEO_Olustur(BlogAd)&".html/Kaydet")
end if
End Sub

Sub NewBlog(Aktiflik)
if BlogAd="" then
	response.Redirect(SiteAdres&"/Yeni/Ad")
elseif BlogIcerik="" then
	response.Redirect(SiteAdres&"/Yeni/Icerik")
else
	Set ObjRs=Server.CreateObject("ADODB.RecordSet"):StrSql="select * from blog":ObjRs.Open StrSql,ObjCon,1,3
	ObjRs.addnew:ObjRs("baslik")=BlogAd:ObjRs("metin")=BlogIcerik:if BlogKategori<>"" then ObjRs("kategori")=BlogKategori: ObjRs("gorunur")=Aktiflik
	ObjRs.update:BlogID=ObjRs("id"):ObjRs.Close:BlogEtiket=split(BlogEtiket&",",",")
	if ubound(BlogEtiket)>0 then
		for i=0 to ubound(BlogEtiket)
			if trim(BlogEtiket(i))>"" then
			StrSql="select * from etiket_ad where etiket='"&trim(BlogEtiket(i))&"'":ObjRs.Open StrSql,ObjCon,1,3
			if ObjRs.eof then ObjRs.addnew:ObjRs("etiket")=trim(BlogEtiket(i)):ObjRs.update
			EtiketID=ObjRs("id"):ObjRs.Close:StrSql="select * from etiket_bulut":ObjRs.Open StrSql,ObjCon,1,3
			ObjRs.addnew:ObjRs("blog_id")=BlogID:ObjRs("etiket_id")=EtiketID:ObjRs.update:ObjRs.close
			end if
		next		
	end if:Set ObjRs=Nothing
	response.Redirect(SiteAdres&"/blog/"&BlogID&"/"&SEO_Olustur(BlogAd)&".html/Yeni")
end if
End Sub
%>